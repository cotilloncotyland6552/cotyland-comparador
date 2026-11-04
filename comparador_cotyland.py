import csv
import json
import os
import sys
import subprocess
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime

APP_TITLE = 'Cotyland | Comparador de Precios y Etiquetas'
APP_VERSION = '7.3.0'


def money(value):
    raw = str(value or '0').strip().replace(' ', '').replace('$', '')
    if ',' in raw:
        raw = raw.replace('.', '').replace(',', '.')
    try:
        return Decimal(raw)
    except InvalidOperation:
        return Decimal('0')


def display_money(value):
    return '$ ' + (f'{value:,.0f}' if value == value.to_integral_value() else f'{value:,.2f}').replace(',', 'X').replace('.', ',').replace('X', '.')


def load_csv(path):
    result = {}
    with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        reader = csv.reader(f)
        for line, row in enumerate(reader, 1):
            # Exportación Cotyland: columnas 0 a 8 son cabeceras REPETIDAS.
            # Los valores reales son columnas 9 (IdArticulo), 10 (Descripcion), 14 (Precio).
            if len(row) >= 15 and row[4].strip().lower() == 'idarticulo' and row[5].strip().lower() == 'descripcion':
                code, name, price = row[9].strip(), row[10].strip(), money(row[14])
            elif line == 1:
                raise ValueError('El archivo no tiene el formato de exportación Cotyland esperado.')
            else:
                continue
            if not code:
                continue
            if code in result:
                raise ValueError(f'Código duplicado en {os.path.basename(path)}: {code} (línea {line})')
            result[code] = dict(code=code, name=name, price=price)
    if not result:
        raise ValueError('No se pudieron leer productos del archivo.')
    return result


class CatalogDB:
    # Persistent database: outside program directory, never replaced by updater.
    def __init__(self):
        self.base = Path(os.environ.get('APPDATA', str(Path.home()))) / 'Cotyland_Comparador'
        self.base.mkdir(parents=True, exist_ok=True)
        self.path = self.base / 'catalogo_tienda.sqlite3'
        self.legacy = self.base / 'productos_en_tienda.json'
        self.conn = sqlite3.connect(str(self.path), timeout=10)
        self.conn.execute('PRAGMA journal_mode=WAL')
        self.conn.execute("CREATE TABLE IF NOT EXISTS catalogo (codigo TEXT PRIMARY KEY, descripcion TEXT NOT NULL DEFAULT '', precio TEXT, en_tienda INTEGER NOT NULL DEFAULT 0, visto_en TEXT, actualizado_en TEXT)")
        self.conn.execute('CREATE TABLE IF NOT EXISTS metadata (clave TEXT PRIMARY KEY, valor TEXT)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS decisiones_manuales (codigo TEXT PRIMARY KEY, en_tienda INTEGER NOT NULL)')
        self.conn.commit()
        # Import legacy selections only once: repeatedly migrating old JSON would recheck unmarked items.
        if not self.conn.execute("SELECT 1 FROM metadata WHERE clave='legacy_migrated'").fetchone():
            self._migrate_legacy()
            with self.conn:
                self.conn.execute("INSERT OR REPLACE INTO metadata(clave,valor) VALUES ('legacy_migrated','1')")

    def _migrate_legacy(self):
        # ADDITIVE migration: never unset store flags.
        if not self.legacy.is_file(): return
        try:
            values = json.loads(self.legacy.read_text(encoding='utf-8-sig'))
            if not isinstance(values, list): return
            with self.conn:
                for code in values:
                    code = str(code).strip()
                    if code:
                        self.conn.execute('INSERT INTO catalogo (codigo,en_tienda) VALUES (?,1) ON CONFLICT(codigo) DO UPDATE SET en_tienda=1', (code,))
        except (OSError, ValueError, sqlite3.Error):
            pass  # Keep any damaged legacy JSON untouched.

    def store_codes(self):
        return {r[0] for r in self.conn.execute('SELECT codigo FROM catalogo WHERE en_tienda=1')}

    def products(self):
        return [dict(code=r[0],name=r[1],price=money(r[2] or '0'))
                for r in self.conn.execute('SELECT codigo,descripcion,precio FROM catalogo WHERE en_tienda=1 ORDER BY descripcion,codigo')]

    def sync_catalog(self, products):
        # Register ALL products, but NEVER auto-mark products as in store.
        stamp = datetime.now().isoformat(timespec='seconds')
        with self.conn:
            self.conn.executemany('INSERT INTO catalogo (codigo,descripcion,precio,visto_en,actualizado_en) VALUES (?,?,?,?,?) ON CONFLICT(codigo) DO UPDATE SET descripcion=CASE WHEN length(excluded.descripcion)>0 THEN excluded.descripcion ELSE catalogo.descripcion END,precio=excluded.precio,visto_en=excluded.visto_en,actualizado_en=excluded.actualizado_en',
                ((x['code'],x['name'],str(x['price']),stamp,stamp) for x in products.values()))

    def set_store(self, codes, value, manual=True):
        codes = sorted({str(x).strip() for x in codes if str(x).strip()})
        with self.conn:
            if manual:
                self.conn.executemany('INSERT INTO decisiones_manuales (codigo,en_tienda) VALUES (?,?) ON CONFLICT(codigo) DO UPDATE SET en_tienda=excluded.en_tienda',
                    ((code,int(value)) for code in codes))
            else:
                # Do not override manual decisions even if a subsequent import differs.
                ignored = {row[0] for row in self.conn.execute('SELECT codigo FROM decisiones_manuales')}
                codes = [c for c in codes if c not in ignored]
            self.conn.executemany('INSERT INTO catalogo (codigo,en_tienda) VALUES (?,?) ON CONFLICT(codigo) DO UPDATE SET en_tienda=excluded.en_tienda',
                 ((code,int(value)) for code in codes))
        self.write_legacy()

    def write_legacy(self):
        # Maintain backward-compatible JSON and prevent lost selections after older reinstall.
        codes = sorted(self.store_codes())
        temp = self.legacy.with_suffix('.json.tmp')
        temp.write_text(json.dumps(codes, ensure_ascii=False), encoding='utf-8')
        os.replace(temp, self.legacy)

    def backup(self, destination):
        with sqlite3.connect(str(destination)) as dest:
            self.conn.backup(dest)



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('1180x740')
        self.minsize(850, 570)
        self.configure(bg='#f3f5f9')
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.search = tk.StringVar()
        self.kind = tk.StringVar(value='Precios modificados')
        self.changes = []
        self.new_items = []
        self.visible = {}
        self.chosen = set()
        self.qty = {}
        self.db = CatalogDB()
        self.in_store = self.db.store_codes()
        self.only_store = tk.BooleanVar(value=True)
        self._build()

    def _build(self):
        style = ttk.Style(self)
        if 'clam' in style.theme_names(): style.theme_use('clam')
        style.configure('Treeview', rowheight=28, font=('Segoe UI', 10), background='white', fieldbackground='white')
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        style.map('Treeview', background=[('selected', '#ffdf69')], foreground=[('selected', '#161616')])
        header = tk.Frame(self, bg='#ffe03a', padx=22, pady=16)
        header.pack(fill='x')
        ttk.Button(header, text='BUSCAR ACTUALIZACIONES', command=self.check_updates).pack(side='right')
        tk.Label(header, text='COTYLAND', font=('Segoe UI', 22, 'bold'), bg='#ffe03a').pack(anchor='w')
        tk.Label(header, text='Comparador de precios y etiquetas · 7 × 3,5 cm', font=('Segoe UI', 11), bg='#ffe03a').pack(anchor='w')
        files = tk.Frame(self, bg='#f3f5f9', padx=18, pady=14)
        files.pack(fill='x')
        self._file_row(files, 'LISTADO ANTERIOR', self.old_path, 0)
        self._file_row(files, 'LISTADO ACTUALIZADO', self.new_path, 1)
        ttk.Button(files, text='COMPARAR LISTADOS', command=self.compare).grid(row=2, column=0, columnspan=3, sticky='ew', pady=(12,0), ipady=6)
        files.columnconfigure(1, weight=1)
        bar = tk.Frame(self, bg='#f3f5f9', padx=18)
        bar.pack(fill='x')
        tk.Label(bar, text='Ver:', bg='#f3f5f9').pack(side='left')
        cb = ttk.Combobox(bar, textvariable=self.kind, values=['Precios modificados','Artículos nuevos','Catálogo en tienda'], state='readonly', width=23)
        cb.pack(side='left', padx=(8,18)); cb.bind('<<ComboboxSelected>>', lambda _: self.refresh())
        self.new_count_label = tk.Label(bar, text='PRODUCTOS NUEVOS: 0', bg='#fff1b8', fg='#332700', font=('Segoe UI', 10, 'bold'), padx=8, pady=5, cursor='hand2')
        self.new_count_label.pack(side='left', padx=(0,12))
        self.new_count_label.bind('<Button-1>', lambda _: self.show_new_products())
        tk.Label(bar, text='Buscar:', bg='#f3f5f9').pack(side='left')
        entry = ttk.Entry(bar, textvariable=self.search)
        entry.pack(side='left', fill='x', expand=True, padx=(8,0))
        ttk.Checkbutton(bar, text='Solo productos en tienda', variable=self.only_store, command=self.refresh).pack(side='left', padx=(12,0))
        self.search.trace_add('write', lambda *_: self.refresh())
        self.status = tk.StringVar(value='Cargá los dos archivos CSV y tocá COMPARAR LISTADOS.')
        tk.Label(self, textvariable=self.status, bg='#f3f5f9', fg='#364152', font=('Segoe UI', 10, 'bold'), anchor='w', padx=20, pady=10).pack(fill='x')
        frame = tk.Frame(self, padx=18, bg='#f3f5f9')
        frame.pack(fill='both', expand=True)
        columns = ('marked', 'store', 'code','name','old','new','qty')
        self.table = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
        for key, title, width, anchor in [('marked','Imprimir',75,'center'),('store','En tienda',75,'center'),('code','Código',125,'w'),('name','Descripción',480,'w'),('old','Precio anterior',140,'e'),('new','Precio nuevo',140,'e'),('qty','Copias',70,'center')]:
            self.table.heading(key, text=title)
            self.table.column(key, width=width, minwidth=60, anchor=anchor, stretch=(key=='name'))
        sc = ttk.Scrollbar(frame, orient='vertical', command=self.table.yview)
        self.table.configure(yscrollcommand=sc.set)
        self.table.pack(side='left', fill='both', expand=True)
        sc.pack(side='right', fill='y')
        self.table.bind('<Button-1>', self.click_checkbox)
        self.table.bind('<Double-1>', self.toggle_selected)
        footer = tk.Frame(self, bg='white', padx=18, pady=9)
        footer.pack(fill='x')
        actions_top = tk.Frame(footer, bg='white')
        actions_top.pack(fill='x', pady=(0,5))
        ttk.Button(actions_top, text='Marcar visibles para imprimir', command=self.mark_visible).pack(side='left', padx=(0,8))
        ttk.Button(actions_top, text='Desmarcar impresión', command=self.unmark).pack(side='left', padx=(0,8))
        ttk.Button(actions_top, text='Copias seleccionadas…', command=self.set_copies).pack(side='left', padx=(0,8))
        ttk.Button(actions_top, text='CREAR ETIQUETAS PDF', command=self.labels).pack(side='right')
        ttk.Button(actions_top, text='Exportar cambios CSV', command=self.export).pack(side='right', padx=(8,8))
        actions_bottom = tk.Frame(footer, bg='white')
        actions_bottom.pack(fill='x')
        ttk.Button(actions_bottom, text='Marcar en tienda', command=lambda: self.set_store(True)).pack(side='left', padx=(0,8))
        ttk.Button(actions_bottom, text='Quitar de tienda', command=lambda: self.set_store(False)).pack(side='left', padx=(0,8))
        ttk.Button(actions_bottom, text='Agregar nuevos a tienda', command=lambda: self.set_store(True)).pack(side='left', padx=(0,8))
        ttk.Button(actions_bottom, text='IMPORTAR LISTA DE TIENDA', command=self.import_store_csv).pack(side='left', padx=(0,8))
        ttk.Button(actions_bottom, text='RESPALDAR BASE', command=self.backup_catalog).pack(side='right')

    def check_updates(self):
        import updater
        from tkinter.simpledialog import askstring
        url = updater.get_update_url()
        if not url:
            url = askstring('Configurar actualizaciones',
                'Para descargar futuras versiones automáticamente, necesitás una dirección HTTPS de actualizaciones.\n'
                'Pegá aquí la URL del archivo version.json publicado en tu servidor o GitHub:', parent=self)
            if not url: return
            try: updater.save_update_url(url)
            except ValueError as exc:
                messagebox.showerror('URL inválida', str(exc)); return
        try:
            result = updater.check_and_install(url)
            if result['status'] == 'current':
                messagebox.showinfo('Actualizado', 'Ya tenés la última versión disponible (' + APP_VERSION + ').')
            else:
                if messagebox.askyesno('Actualización instalada',
                        'Se instaló la versión ' + result['version'] + '.\n\n¿Reiniciar el programa ahora?'):
                    self.destroy()
                    subprocess.Popen([str(Path(sys.executable).with_name('pythonw.exe')) if sys.platform == 'win32' and Path(sys.executable).with_name('pythonw.exe').exists() else sys.executable, str(Path(__file__).resolve())], cwd=str(Path(__file__).resolve().parent))
        except Exception as exc:
            messagebox.showerror('No se pudo actualizar', str(exc) +
                '\n\nPodés cambiar la dirección editando actualizaciones.json junto al programa.')

    def _file_row(self, parent, label, var, row):
        tk.Label(parent, text=label, bg='#f3f5f9', font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, padx=(0,12), pady=6, sticky='w')
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky='ew', pady=6)
        ttk.Button(parent, text='Examinar…', command=lambda: self.browse(var)).grid(row=row, column=2, padx=(12,0), pady=6)

    def browse(self, var):
        path = filedialog.askopenfilename(title='Seleccionar listado CSV', filetypes=[('CSV', '*.csv'), ('Todos los archivos','*.*')])
        if path: var.set(path)

    def compare(self):
        if not self.old_path.get() or not self.new_path.get():
            messagebox.showwarning('Faltan archivos', 'Seleccioná los dos listados.'); return
        try:
            self.configure(cursor='watch'); self.update_idletasks()
            before = load_csv(self.old_path.get())
            after = load_csv(self.new_path.get())
            self.changes = [dict(new=entry, old=before[code]['price']) for code, entry in after.items() if code in before and entry['price'] != before[code]['price']]
            self.new_items = [dict(new=entry, old=None) for code, entry in after.items() if code not in before]
            self.db.sync_catalog(after)
            self.in_store = self.db.store_codes()
            self.changes.sort(key=lambda item: item['new']['name'].casefold())
            self.new_items.sort(key=lambda item: item['new']['name'].casefold())
            self.new_count_label.config(text=f'PRODUCTOS NUEVOS: {len(self.new_items):,}'.replace(',', '.'))
            self.chosen.clear(); self.qty.clear()
            self.kind.set('Precios modificados'); self.search.set('')
            self.summary = f'{len(before):,} anteriores · {len(after):,} actuales · {len(self.changes):,} precios cambiados · {len(self.new_items):,} productos nuevos'.replace(',', '.')
            self.refresh()
        except Exception as e:
            messagebox.showerror('Error al leer los archivos', str(e))
        finally:
            self.configure(cursor='')

    def current(self):
        if self.kind.get() == 'Precios modificados': return self.changes
        if self.kind.get() == 'Artículos nuevos': return self.new_items
        return [dict(new=p,old=None) for p in self.db.products()]

    def refresh(self):
        for item in self.table.get_children(): self.table.delete(item)
        q = self.search.get().strip().casefold()
        self.visible = {}
        for item in self.current():
            entry = item['new']; code = entry['code']
            if q and q not in code.casefold() and q not in entry['name'].casefold(): continue
            if self.kind.get() == 'Precios modificados' and self.only_store.get() and code not in self.in_store: continue
            # New items always visible so user can decide which to add.
            self.visible[code] = item
            self.table.insert('', 'end', iid=code, values=('☑' if code in self.chosen else '□', '☑' if code in self.in_store else '□', code, entry['name'], display_money(item['old']) if item['old'] is not None else '—', display_money(entry['price']), self.qty.get(code,1)))
        self.status.set((getattr(self,'summary','Sin comparación') + f'   |   Mostrando {len(self.visible)}   |   En tienda {len(self.in_store)}   |   Para imprimir {len(self.chosen)}'))

    def save_store(self, codes, value):
        try:
            self.db.set_store(codes, value)
            self.in_store = self.db.store_codes()
        except (OSError, sqlite3.Error) as exc:
            messagebox.showerror('No se pudo guardar la base de datos', str(exc))

    def backup_catalog(self):
        path = filedialog.asksaveasfilename(title='Guardar copia de seguridad del catálogo',
            defaultextension='.sqlite3', initialfile='Cotyland_catalogo_respaldo.sqlite3',
            filetypes=[('Base de datos SQLite','*.sqlite3')])
        if not path: return
        try:
            self.db.backup(path)
            messagebox.showinfo('Respaldo guardado', f'Se guardó una copia del catálogo en:\n{path}')
        except (OSError,sqlite3.Error) as exc:
            messagebox.showerror('Error de respaldo', str(exc))

    def show_new_products(self):
        self.kind.set('Artículos nuevos')
        self.refresh()

    def import_store_csv(self):
        """Import store codes without clearing the old selection. Accepts Cotyland's
        repeated-header 16-column CSV, simple CSV with Codigo column, or a stock CSV.
        An active flag or a price is never automatically interpreted as stock.
        """
        path = filedialog.askopenfilename(title='Elegir lista para marcar artículos en tienda',
            filetypes=[('Archivos CSV','*.csv'),('Todos los archivos','*.*')])
        if not path: return
        try:
            self.configure(cursor='watch'); self.update_idletasks()
            rows, origin, stock_present = self._read_store_import(path)
            if not rows: raise ValueError('No se encontraron códigos válidos.')
        except (OSError, ValueError, csv.Error) as exc:
            messagebox.showerror('No se pudo leer el listado', str(exc)); return
        finally:
            self.configure(cursor='')

        # For stock files, any non-zero quantity (including negative) qualifies.
        # For an active-only export, require an explicit opt-in to provisional marking.
        with_stock = sum(1 for p in rows.values() if p.get('stock') is not None and p['stock'] != 0)
        priced = sum(1 for p in rows.values() if p.get('price') is not None and p['price'] > 0)
        zero_stock = sum(1 for p in rows.values() if p.get('stock') == 0)
        dialog = tk.Toplevel(self)
        dialog.title('Confirmar importación · Cotyland')
        dialog.geometry('720x455')
        dialog.minsize(640,400)
        dialog.transient(self); dialog.grab_set()
        container = ttk.Frame(dialog, padding=20); container.pack(fill='both',expand=True)
        ttk.Label(container,text=f'Archivo: {Path(path).name}',font=('Segoe UI',12,'bold')).pack(anchor='w',pady=(0,8))
        ttk.Label(container,text=f'Códigos encontrados: {len(rows):,} | Con precio positivo: {priced:,}' +
                  (f' | Stock distinto de cero: {with_stock:,} | Stock cero: {zero_stock:,}' if stock_present else ''),
                  font=('Segoe UI',10)).pack(anchor='w',pady=(0,9))
        note = ('Se detectó una columna de existencias. Se pueden marcar automáticamente los '
                'productos con stock positivo o negativo; los de stock cero NO se desmarcan.') if stock_present else (
                'ATENCIÓN: este archivo NO tiene una columna verificable de stock. '
                'ACTIVO o PRECIO > 0 NO significa necesariamente que el producto esté físicamente en tienda. '
                'Podés usarlo como selección preliminar, pero conviene revisar o conseguir el reporte de existencias.')
        ttk.Label(container,text=note,wraplength=665,justify='left',foreground='#a53b17').pack(anchor='w',pady=(0,12))
        mode = tk.StringVar(value='stock' if stock_present else 'none')
        if stock_present:
            ttk.Radiobutton(container,variable=mode,value='stock',text=f'Marcar solo stock distinto de 0 (incluye negativos): {with_stock:,}').pack(anchor='w',pady=3)
        ttk.Radiobutton(container,variable=mode,value='none',text='Solo cargar datos de artículos, SIN marcar ninguno en tienda').pack(anchor='w',pady=3)
        ttk.Radiobutton(container,variable=mode,value='priced',text=f'Marcar como selección PROVISIONAL los {priced:,} artículos con precio mayor a cero').pack(anchor='w',pady=3)
        ttk.Radiobutton(container,variable=mode,value='all',text=f'Marcar TODOS los {len(rows):,} códigos del archivo (solo si realmente son de la tienda)').pack(anchor='w',pady=3)
        ttk.Label(container,text='Las marcas existentes se conservan y las decisiones manuales tienen prioridad sobre nuevas importaciones.',wraplength=665).pack(anchor='w',pady=(12,4))
        def apply():
            selection = mode.get()
            codes = ([p['code'] for p in rows.values() if p.get('stock') is not None and p['stock'] != 0]
                     if selection == 'stock' else
                     [p['code'] for p in rows.values() if p.get('price') is not None and p['price'] > 0]
                     if selection == 'priced' else list(rows) if selection == 'all' else [])
            if selection in ('priced','all'):
                if not messagebox.askyesno('Confirmación especial',
                    f'Vas a marcar {len(codes):,} artículos como «En tienda» SIN verificar el stock.\n\n'
                    '¿Confirmás esta selección? Se sumará a lo ya marcado.', parent=dialog):
                    return
            elif selection == 'stock':
                if not messagebox.askyesno('Confirmar stock',
                    f'Se van a marcar {len(codes):,} artículos con stock distinto de cero.\n'
                    'No se quitará ninguna marca existente. ¿Continuar?', parent=dialog):
                    return
            try:
                self.db.sync_catalog({p['code']:dict(code=p['code'],name=p['name'],price=p.get('price') or Decimal(0)) for p in rows.values()})
                if codes: self.db.set_store(codes, True, manual=False)
                self.in_store = self.db.store_codes()
                dialog.destroy()
                self.refresh()
                messagebox.showinfo('Importación realizada',
                    f'Artículos importados a la base: {len(rows):,}\n'
                    f'Códigos marcados en esta operación: {len(codes):,}\n'
                    f'Total guardado «En tienda»: {len(self.in_store):,}\n\n'
                    'Las marcas seguirán guardadas al actualizar el comparador.')
            except (OSError,sqlite3.Error) as exc:
                messagebox.showerror('Error al guardar la base',str(exc),parent=dialog)
        buttons=ttk.Frame(container); buttons.pack(anchor='e',side='bottom',pady=(10,0))
        ttk.Button(buttons,text='Cancelar',command=dialog.destroy).pack(side='right',padx=6)
        ttk.Button(buttons,text='IMPORTAR',command=apply).pack(side='right')

    @staticmethod
    def _read_store_import(path):
        """Returns dictionary code -> {code,name,price,stock} and whether stock exists."""
        result={}; stock_found=False
        with open(path,'r',encoding='utf-8-sig',errors='replace',newline='') as f:
            sample=f.readline(); f.seek(0)
            reader=csv.reader(f,delimiter=';' if sample.count(';') > sample.count(',') else ',')
            first=next(reader, None)
            if first is None: return {},'empty',False
            # Standard export: repeated label columns + real values at 9,10,14.
            cotyland=(len(first)>=15 and first[4].strip().lower()=='idarticulo'
                      and first[5].strip().lower()=='descripcion')
            if cotyland:
                def parse_row(row):
                    if len(row)<15 or row[4].strip().lower()!='idarticulo': return
                    code=row[9].strip()
                    if not code:return
                    result[code]={'code':code,'name':row[10].strip(), 'price':money(row[14]),'stock':None}
                parse_row(first)
                for row in reader:parse_row(row)
                return result,'Cotyland',False
            names=[c.strip().casefold().replace(' ','').replace('_','') for c in first]
            code_names={'idarticulo','codigo','código','sku','codarticulo','codigodearticulo'}
            code_idx=next((i for i,c in enumerate(names) if c in code_names),None)
            if code_idx is None: raise ValueError('No se identificó la columna de código. Usá IdArticulo, Codigo o SKU.')
            stock_names={'stock','existencias','existencia','stockactual','cantidadstock','cantidadenstock','saldo','stockdisponible'}
            stock_idx=next((i for i,c in enumerate(names) if c in stock_names),None)
            stock_found=stock_idx is not None
            price_idx=next((i for i,c in enumerate(names) if c in ('precio','precioventa','preciolista')),None)
            name_idx=next((i for i,c in enumerate(names) if c in ('descripcion','descripción','nombre','producto')),None)
            for line,row in enumerate(reader,2):
                if len(row)<=code_idx:continue
                code=row[code_idx].strip()
                if not code:continue
                p=money(row[price_idx]) if price_idx is not None and len(row)>price_idx else None
                qty=None
                if stock_found and len(row)>stock_idx:
                    raw=row[stock_idx].strip().replace(' ','')
                    if raw:
                        try:
                            if ',' in raw:raw=raw.replace('.','').replace(',','.')
                            qty=Decimal(raw)
                        except InvalidOperation:
                            raise ValueError(f'Stock inválido para el artículo {code} en fila {line}: {row[stock_idx]}')
                result[code]={'code':code,'name':row[name_idx].strip() if name_idx is not None and len(row)>name_idx else '',
                              'price':p,'stock':qty}
        return result,'simple',stock_found

    def click_checkbox(self, event):
        row = self.table.identify_row(event.y)
        col = self.table.identify_column(event.x)
        if not row or self.table.identify_region(event.x, event.y) != 'cell': return
        if col == '#1':
            self.chosen.symmetric_difference_update({row})
        elif col == '#2':
            self.save_store([row], row not in self.in_store)
        else: return
        self.refresh()

    def set_store(self, value):
        codes = set(self.table.selection())
        if not codes:
            messagebox.showinfo('Seleccioná filas', 'Seleccioná una o más filas con un clic (Ctrl o Shift para varias) y después usá este botón. También podés pulsar directamente la casilla En tienda.'); return
        self.save_store(codes, value); self.refresh()

    def toggle_selected(self, event):
        row = self.table.identify_row(event.y)
        if not row or self.table.identify_column(event.x) in ('#1','#2'): return
        if row in self.chosen: self.chosen.remove(row)
        else: self.chosen.add(row)
        self.refresh()

    def mark_visible(self):
        self.chosen.update(self.visible)
        self.refresh()

    def unmark(self):
        self.chosen.clear(); self.refresh()

    def set_copies(self):
        from tkinter.simpledialog import askinteger
        number = askinteger('Cantidad de etiquetas', '¿Cuántas etiquetas por producto seleccionado?', initialvalue=1, minvalue=1, maxvalue=1000, parent=self)
        if number is not None:
            for code in self.chosen: self.qty[code] = number
            self.refresh()

    def export(self):
        if not self.changes and not self.new_items:
            messagebox.showinfo('Sin datos', 'Primero compará los listados.'); return
        path = filedialog.asksaveasfilename(defaultextension='.csv', initialfile='cotyland_cambios.csv', filetypes=[('CSV','*.csv')])
        if not path: return
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f, delimiter=';')
            w.writerow(['Tipo','IdArticulo','Descripcion','Precio anterior','Precio nuevo'])
            for type_name, rows in [('Cambio de precio',self.changes),('Artículo nuevo',self.new_items)]:
                for item in rows:
                    w.writerow([type_name,item['new']['code'],item['new']['name'],str(item['old']) if item['old'] is not None else '',str(item['new']['price'])])
        messagebox.showinfo('Exportación completa', f'Archivo guardado en:\n{path}')

    def labels(self):
        if not self.chosen:
            messagebox.showwarning('Seleccioná productos', 'Marcá los productos que quieras imprimir (doble clic o Marcar visibles).'); return
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import mm
            from reportlab.pdfbase.pdfmetrics import stringWidth
        except ImportError:
            messagebox.showerror('Falta dependencia', 'Instalá ReportLab con: py -m pip install reportlab'); return
        # Evitar impresiones accidentales de decenas de miles de páginas.
        rows = {x['new']['code']:x for x in self.changes + self.new_items}
        selected = [rows[k] for k in self.chosen if k in rows]
        copies = sum(self.qty.get(x['new']['code'], 1) for x in selected)
        if copies > 5000 and not messagebox.askyesno('Muchas etiquetas',f'Se van a generar {copies} etiquetas. ¿Continuar?'): return
        path = filedialog.asksaveasfilename(defaultextension='.pdf', initialfile='etiquetas_cotyland_A4_24.pdf', filetypes=[('PDF','*.pdf')])
        if not path: return
        # A4 de 210 x 297 mm: 3 columnas x 8 filas de 70 x 35 mm.
        # Sin margen horizontal ni espacios entre etiquetas. Sobra una franja
        # vertical inevitable de 17 mm, que se reparte arriba y abajo.
        page_w, page_h = 210*mm, 297*mm
        w, h = 70*mm, 35*mm
        columns, per_page = 3, 24
        bottom = (page_h - 8*h)/2
        c = canvas.Canvas(path, pagesize=(page_w, page_h), pageCompression=1)
        generated_date = datetime.now().strftime('%d/%m/%Y')

        def clip(text, font, size, width):
            text = ' '.join(str(text).split())
            if stringWidth(text, font, size) <= width:
                return text
            while text and stringWidth(text+'…', font, size) > width:
                text = text[:-1]
            return text.rstrip() + '…'

        def description_lines(txt):
            # Hasta 2 líneas en el ancho útil. Evita cortar de golpe los nombres largos.
            words = ' '.join(str(txt or 'ARTÍCULO SIN DESCRIPCIÓN').split()).split()
            lines = ['']
            for word in words:
                try_line = (lines[-1]+' '+word).strip()
                if stringWidth(try_line, 'Helvetica-Bold', 8.2) <= w-5*mm:
                    lines[-1] = try_line
                elif len(lines) < 2:
                    lines.append(word)
                else:
                    lines[-1] = clip(lines[-1] + ' ' + word, 'Helvetica-Bold', 8.2, w-5*mm)
                    break
            return lines

        index = 0
        for item in sorted(selected, key=lambda x: x['new']['name'].casefold()):
            product = item['new']
            for _ in range(self.qty.get(product['code'], 1)):
                if index and index % per_page == 0:
                    c.showPage()
                slot = index % per_page
                col, row = slot % columns, slot // columns
                x = col*w
                y = page_h-bottom-(row+1)*h
                c.setLineWidth(0.27)
                c.setStrokeColorRGB(.65,.65,.65)
                c.rect(x,y,w,h,stroke=1,fill=0)
                c.setFillColorRGB(0,0,0)

                lines = description_lines(product['name'])
                c.setFont('Helvetica-Bold', 8.2)
                for n, line in enumerate(lines):
                    c.drawCentredString(x+w/2, y+h-(3.8+n*3.3)*mm, line)

                # Precio protagonista: sin barras, centrado y siempre dentro de la etiqueta.
                # El tamaño máximo real depende del ancho de los 70 mm.
                price = display_money(product['price'])
                max_price_width = w - 4*mm
                desired_size = 48.0
                size = min(desired_size, desired_size * max_price_width /
                           max(stringWidth(price, 'Helvetica-Bold', desired_size), 1))
                c.setFont('Helvetica-Bold', size)
                # Línea de precio entre la descripción y el código numérico.
                c.drawCentredString(x+w/2, y+12.0*mm, price)

                # Código exclusivamente numérico, sin barras ni fuente de barcode.
                identifier = str(product['code']).strip()
                c.setFont('Helvetica-Bold', 10.0)
                c.drawCentredString(x+w/2, y+5.1*mm,
                    clip('Cód. art.: ' + identifier, 'Helvetica-Bold', 10.0, w-4*mm))
                c.setFont('Helvetica', 7.0)
                c.drawCentredString(x+w/2, y+1.9*mm, 'Actualizado: ' + generated_date)
                index += 1
        c.save()
        if messagebox.askyesno('Etiquetas generadas',f'Guardadas {copies} etiquetas en hojas A4 (24 etiquetas por página, 70 × 35 mm cada una).\n\n¿Abrir el PDF para imprimir?'):
            try:
                if sys.platform == 'win32': os.startfile(path)
                elif sys.platform == 'darwin': subprocess.Popen(['open',path])
                else: subprocess.Popen(['xdg-open',path])
            except Exception as e: messagebox.showinfo('Archivo generado',f'Encontralo en: {path}\n{e}')

if __name__=='__main__':
    App().mainloop()
