import csv
import json
import os
import sys
import subprocess
import sqlite3
import hashlib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime

APP_TITLE = 'Cotyland | Comparador de Precios y Etiquetas'
APP_VERSION = '7.5.3'


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


# Configuración local: no se publica en GitHub ni se sincroniza la base SQLite.
DRIVE_DEFAULT = Path(r'G:\Mi unidad\Comparador NUEVO de precios\Listas de Precios')
DRIVE_SETTINGS = Path(os.environ.get('APPDATA', str(Path.home()))) / 'Cotyland_Comparador' / 'listas_drive.json'


def drive_folder_saved():
    try:
        value = json.loads(DRIVE_SETTINGS.read_text(encoding='utf-8')).get('carpeta', '')
        return Path(value) if value else DRIVE_DEFAULT
    except (OSError, ValueError, TypeError, AttributeError):
        return DRIVE_DEFAULT


def save_drive_folder(folder):
    DRIVE_SETTINGS.parent.mkdir(parents=True, exist_ok=True)
    DRIVE_SETTINGS.write_text(json.dumps({'carpeta':str(folder)}, ensure_ascii=False, indent=2), encoding='utf-8')


def find_drive_pair(folder):
    # Windows es case-insensitive; permitir también Linux o entornos de prueba.
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError('La carpeta de Drive no está accesible: '+str(folder))
    entries = {p.name.casefold(): p for p in folder.iterdir() if p.is_file()}
    old = entries.get('viejo.csv')
    new = entries.get('nuevo.csv')
    if not old or not new:
        raise ValueError('En la carpeta deben estar los dos archivos: viejo.csv y nuevo.csv.')
    for p in (old, new):
        if p.stat().st_size < 100:
            raise ValueError(f'El archivo {p.name} parece vacío o incompleto.')
    if old.stat().st_mtime > new.stat().st_mtime + 2:
        raise ValueError('La fecha de viejo.csv es posterior a la de nuevo.csv. Revisá que sean del mismo cambio de precios.')
    return old, new


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
        self.conn.execute('CREATE TABLE IF NOT EXISTS precios_manuales (codigo TEXT PRIMARY KEY, precio TEXT NOT NULL)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS codigos_alternativos (alternativo TEXT PRIMARY KEY, codigo TEXT NOT NULL)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS decisiones_impresion (lote TEXT NOT NULL, codigo TEXT NOT NULL, imprimir INTEGER NOT NULL, PRIMARY KEY(lote,codigo))')
        self.conn.execute('CREATE TABLE IF NOT EXISTS preferencias_impresion (codigo TEXT NOT NULL, precio TEXT NOT NULL, imprimir INTEGER NOT NULL, PRIMARY KEY(codigo,precio))')
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
            self.conn.executemany('INSERT INTO catalogo (codigo,descripcion,precio,visto_en,actualizado_en) VALUES (?,?,?,?,?) ON CONFLICT(codigo) DO UPDATE SET descripcion=CASE WHEN length(excluded.descripcion)>0 THEN excluded.descripcion ELSE catalogo.descripcion END,precio=CASE WHEN EXISTS(SELECT 1 FROM precios_manuales WHERE codigo=catalogo.codigo) THEN catalogo.precio ELSE excluded.precio END,visto_en=excluded.visto_en,actualizado_en=excluded.actualizado_en',
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

    def print_decisions(self, lot):
        if not lot:
            return {}
        return {code: bool(value) for code, value in
                self.conn.execute('SELECT codigo,imprimir FROM decisiones_impresion WHERE lote=?', (lot,))}

    def set_print_decisions(self, lot, codes, value):
        if not lot:
            return
        codes = {str(code).strip() for code in codes if str(code).strip()}
        with self.conn:
            self.conn.executemany(
                'INSERT INTO decisiones_impresion(lote,codigo,imprimir) VALUES (?,?,?) '
                'ON CONFLICT(lote,codigo) DO UPDATE SET imprimir=excluded.imprimir',
                ((lot, code, int(bool(value))) for code in codes))

    def print_preferences(self, items):
        """Saved print decisions by article + target price, independent of CSV revisions."""
        if not items:
            return {}
        codes = [item['new']['code'] for item in items]
        placeholders = ','.join('?' for _ in codes)
        saved = self.conn.execute(
            f'SELECT codigo,precio,imprimir FROM preferencias_impresion WHERE codigo IN ({placeholders})', codes).fetchall()
        by_key = {(code, price): bool(flag) for code, price, flag in saved}
        return {item['new']['code']: by_key[(item['new']['code'], str(item['new']['price']))]
                for item in items if (item['new']['code'], str(item['new']['price'])) in by_key}

    def set_print_preferences(self, rows):
        """rows: iterable of (article code, updated price, print flag)."""
        with self.conn:
            self.conn.executemany(
                'INSERT INTO preferencias_impresion(codigo,precio,imprimir) VALUES (?,?,?) '
                'ON CONFLICT(codigo,precio) DO UPDATE SET imprimir=excluded.imprimir',
                [(str(code),str(price),int(bool(enabled))) for code,price,enabled in rows])

    def link_barcode(self, barcode, code):
        barcode, code = str(barcode).strip(), str(code).strip()
        if not barcode or not code:
            raise ValueError('Ingresá un código de barras y un código interno.')
        if not self.get_product(code):
            raise ValueError(f'El código interno {code} no existe en el catálogo. Primero cargalo.')
        if barcode == code:
            return
        existing = self.conn.execute('SELECT codigo FROM codigos_alternativos WHERE alternativo=?', (barcode,)).fetchone()
        if existing and existing[0] != code:
            raise ValueError(f'El código de barras {barcode} ya está vinculado al artículo {existing[0]}.')
        with self.conn:
            self.conn.execute('INSERT INTO codigos_alternativos(alternativo,codigo) VALUES (?,?) '
                              'ON CONFLICT(alternativo) DO UPDATE SET codigo=excluded.codigo', (barcode, code))

    def import_barcodes(self, rows):
        """Validate whole batch first. Never replace a mapping for another item."""
        mapping = {}
        for code, barcode in rows:
            code, barcode = str(code).strip(), str(barcode).strip()
            if not barcode or not code or code == barcode:
                continue
            if barcode in mapping and mapping[barcode] != code:
                raise ValueError(f'Código de barras duplicado: {barcode} para dos artículos.')
            mapping[barcode] = code
        available = {r[0] for r in self.conn.execute('SELECT codigo FROM catalogo')}
        current = dict(self.conn.execute('SELECT alternativo,codigo FROM codigos_alternativos'))
        missing = {c for c in mapping.values() if c not in available}
        if missing:
            raise ValueError(f'Hay {len(missing)} códigos internos que no figuran en la base. Ejemplo: {sorted(missing)[0]}')
        conflict = [(b,c,current[b]) for b,c in mapping.items() if b in current and current[b] != c]
        if conflict:
            b,c,old = conflict[0]
            raise ValueError(f'Conflicto: el código de barras {b} ya está asignado a {old}, no a {c}.')
        with self.conn:
            self.conn.executemany('INSERT INTO codigos_alternativos(alternativo,codigo) VALUES (?,?) '
                                  'ON CONFLICT(alternativo) DO UPDATE SET codigo=excluded.codigo', mapping.items())
        return len(mapping)

    def write_legacy(self):
        # Maintain backward-compatible JSON and prevent lost selections after older reinstall.
        codes = sorted(self.store_codes())
        temp = self.legacy.with_suffix('.json.tmp')
        temp.write_text(json.dumps(codes, ensure_ascii=False), encoding='utf-8')
        os.replace(temp, self.legacy)

    def find_products(self, term, limit=75):
        term = str(term).strip()
        if not term: return []
        # Scanner sends the same data as a keyboard, normally followed by Enter.
        norm = '%' + term.replace('\\','\\\\').replace('%','\\%').replace('_','\\_') + '%'
        return self.conn.execute(
            "SELECT catalogo.codigo,catalogo.descripcion,catalogo.precio,catalogo.en_tienda FROM catalogo "
            "LEFT JOIN codigos_alternativos a ON a.codigo=catalogo.codigo "
            "WHERE catalogo.codigo LIKE ? ESCAPE '\\' OR catalogo.descripcion LIKE ? ESCAPE '\\' "
            "OR a.alternativo LIKE ? ESCAPE '\\' "
            "ORDER BY CASE WHEN catalogo.codigo=? OR a.alternativo=? THEN 0 ELSE 1 END, catalogo.descripcion LIMIT ?",
            (norm,norm,norm,term,term,limit)).fetchall()

    def save_product(self, code, name, price, in_store=True, barcode=''):
        code = str(code).strip()
        if not code: raise ValueError('El código no puede estar vacío.')
        price = money(price)
        if price < 0: raise ValueError('El precio debe ser mayor o igual a cero.')
        if barcode and barcode.strip() != code:
            prior = self.conn.execute('SELECT codigo FROM codigos_alternativos WHERE alternativo=?', (barcode.strip(),)).fetchone()
            if prior and prior[0] != code:
                raise ValueError(f'El código de barras ya pertenece al artículo {prior[0]}.')
        stamp = datetime.now().isoformat(timespec='seconds')
        with self.conn:
            self.conn.execute("INSERT INTO catalogo(codigo,descripcion,precio,en_tienda,actualizado_en) "
                "VALUES(?,?,?,?,?) ON CONFLICT(codigo) DO UPDATE SET "
                "descripcion=excluded.descripcion,precio=excluded.precio,actualizado_en=excluded.actualizado_en",
                (code, name.strip(), str(price), int(bool(in_store)),stamp))
            self.conn.execute('INSERT INTO precios_manuales(codigo,precio) VALUES (?,?) ON CONFLICT(codigo) DO UPDATE SET precio=excluded.precio', (code,str(price)))
            if barcode and barcode.strip() != code:
                self.conn.execute('INSERT INTO codigos_alternativos(alternativo,codigo) VALUES (?,?) ON CONFLICT(alternativo) DO UPDATE SET codigo=excluded.codigo',(barcode.strip(),code))
        if in_store:
            self.set_store([code], True, manual=True)

    def get_manual_price(self, code):
        r=self.conn.execute('SELECT precio FROM precios_manuales WHERE codigo=?',(code,)).fetchone()
        return money(r[0]) if r else None

    def get_product(self, code):
        r=self.conn.execute('SELECT codigo,descripcion,precio,en_tienda FROM catalogo WHERE codigo=?',(code,)).fetchone()
        return dict(code=r[0],name=r[1],price=money(r[2] or '0'),store=bool(r[3])) if r else None

    def backup(self, destination):
        with sqlite3.connect(str(destination)) as dest:
            self.conn.backup(dest)



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('1280x800')
        self.minsize(980, 620)
        self.configure(bg='#f3f5f9')
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.drive_folder = drive_folder_saved()
        self.search = tk.StringVar()
        self.kind = tk.StringVar(value='Precios modificados')
        self.changes = []
        self.new_items = []
        self.visible = {}
        self.chosen = set()
        self.current_lot = None
        self.qty = {}
        self.db = CatalogDB()
        self.in_store = self.db.store_codes()
        self.only_store = tk.BooleanVar(value=True)
        self._build()

    def _build(self):
        style = ttk.Style(self)
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        style.configure('Treeview', rowheight=26, font=('Segoe UI', 9), background='white', fieldbackground='white')
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))
        style.map('Treeview', background=[('selected', '#ffdf69')], foreground=[('selected', '#161616')])
        style.configure('Compact.TButton', padding=(8, 4), font=('Segoe UI', 9))
        style.configure('Icon.TButton', padding=(6, 4), font=('Segoe UI Symbol', 9))
        style.configure('Primary.TButton', padding=(10, 5), font=('Segoe UI', 9, 'bold'))

        header = tk.Frame(self, bg='#ffe03a', padx=20, pady=10)
        header.pack(fill='x')
        ttk.Button(header, text='BUSCAR ACTUALIZACIONES', command=self.check_updates, style='Compact.TButton').pack(side='right')
        tk.Label(header, text='COTYLAND', font=('Segoe UI', 20, 'bold'), bg='#ffe03a').pack(anchor='w')
        tk.Label(header, text='Comparador de precios y etiquetas · 7 × 3,5 cm', font=('Segoe UI', 10), bg='#ffe03a').pack(anchor='w')

        files = tk.Frame(self, bg='#f3f5f9', padx=18, pady=10)
        files.pack(fill='x')
        self._file_row(files, 'LISTADO ANTERIOR', self.old_path, 0)
        self._file_row(files, 'LISTADO ACTUALIZADO', self.new_path, 1)
        files.columnconfigure(1, weight=1)
        actions = tk.Frame(files, bg='#f3f5f9')
        actions.grid(row=2, column=0, columnspan=3, sticky='ew', pady=(8, 0))
        ttk.Button(actions, text='COMPARAR LISTADOS', command=self.compare, style='Compact.TButton').pack(side='left', fill='x', expand=True)
        ttk.Button(actions, text='☁ CARGAR DESDE DRIVE', command=self.load_from_drive, style='Compact.TButton').pack(side='left', padx=(8, 0))
        ttk.Button(actions, text='📁', command=self.choose_drive_folder, style='Icon.TButton').pack(side='left', padx=(6, 0))

        bar = tk.Frame(self, bg='#f3f5f9', padx=18, pady=2)
        bar.pack(fill='x')
        tk.Label(bar, text='Ver:', bg='#f3f5f9').pack(side='left')
        cb = ttk.Combobox(bar, textvariable=self.kind, values=['Precios modificados', 'Artículos nuevos', 'Catálogo en tienda'], state='readonly', width=20)
        cb.pack(side='left', padx=(6, 10))
        cb.bind('<<ComboboxSelected>>', lambda _: self.refresh())
        ttk.Button(bar, text='Catálogo inicial', command=self.import_store_csv, style='Compact.TButton').pack(side='left', padx=(0, 10))
        self.new_count_label = tk.Label(bar, text='PRODUCTOS NUEVOS: 0', bg='#fff1b8', fg='#332700', font=('Segoe UI', 9, 'bold'), padx=8, pady=3, cursor='hand2')
        self.new_count_label.pack(side='left', padx=(0, 12))
        self.new_count_label.bind('<Button-1>', lambda _: self.show_new_products())
        tk.Label(bar, text='Buscar:', bg='#f3f5f9').pack(side='left')
        entry = ttk.Entry(bar, textvariable=self.search)
        entry.pack(side='left', fill='x', expand=True, padx=(6, 8))
        ttk.Checkbutton(bar, text='Solo productos en tienda', variable=self.only_store, command=self.refresh).pack(side='left')
        self.search.trace_add('write', lambda *_: self.refresh())

        self.status = tk.StringVar(value='Cargá los dos archivos CSV y tocá COMPARAR LISTADOS.')
        tk.Label(self, textvariable=self.status, bg='#f3f5f9', fg='#364152', font=('Segoe UI', 9, 'bold'), anchor='w', padx=18, pady=5).pack(fill='x')

        tools = tk.Frame(self, bg='#f3f5f9', padx=18, pady=2)
        tools.pack(fill='x')
        left_tools = tk.Frame(tools, bg='#f3f5f9')
        left_tools.pack(side='left', fill='x', expand=True)
        ttk.Button(left_tools, text='Marcar visibles', command=self.mark_visible, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Desmarcar', command=self.unmark, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Marcar tienda', command=lambda: self.set_store(True), style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Quitar tienda', command=lambda: self.set_store(False), style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Agregar nuevos', command=lambda: self.set_store(True), style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Importar lista', command=self.import_store_csv, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Copias…', command=self.set_copies, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='🔎 Buscar / escanear', command=self.scan_product, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Importar barras', command=self.import_barcode_csv, style='Compact.TButton').pack(side='left', padx=(0, 6))

        right_tools = tk.Frame(tools, bg='#f3f5f9')
        right_tools.pack(side='right')
        ttk.Button(right_tools, text='🗄 Respaldo', command=self.backup_catalog, style='Compact.TButton').pack(side='right', padx=(6, 0))
        ttk.Button(right_tools, text='▌║█║ Etiquetas', command=self.labels, style='Primary.TButton').pack(side='right', padx=(6, 0))
        ttk.Button(right_tools, text='⤓ Exportar', command=self.export, style='Compact.TButton').pack(side='right', padx=(6, 0))

        frame = tk.Frame(self, padx=18, pady=8, bg='#f3f5f9')
        frame.pack(fill='both', expand=True)
        columns = ('marked', 'store', 'code', 'name', 'old', 'new', 'qty')
        self.table = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
        specs = [
            ('marked', 'Imprimir', 72, 'center'),
            ('store', 'En tienda', 72, 'center'),
            ('code', 'Código', 105, 'w'),
            ('name', 'Descripción', 560, 'w'),
            ('old', 'Precio anterior', 120, 'e'),
            ('new', 'Precio nuevo', 120, 'e'),
            ('qty', 'Copias', 60, 'center'),
        ]
        for key, title, width, anchor in specs:
            self.table.heading(key, text=title)
            self.table.column(key, width=width, minwidth=52, anchor=anchor, stretch=(key == 'name'))
        sc_y = ttk.Scrollbar(frame, orient='vertical', command=self.table.yview)
        sc_x = ttk.Scrollbar(frame, orient='horizontal', command=self.table.xview)
        self.table.configure(yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)
        self.table.grid(row=0, column=0, sticky='nsew')
        sc_y.grid(row=0, column=1, sticky='ns')
        sc_x.grid(row=1, column=0, sticky='ew')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.table.bind('<Button-1>', self.click_checkbox)
        self.table.bind('<Double-1>', self.toggle_selected)
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
        tk.Label(parent, text=label, bg='#f3f5f9', font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, padx=(0,10), pady=4, sticky='w')
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky='ew', pady=4)
        ttk.Button(parent, text='Examinar', command=lambda: self.browse(var), style='Compact.TButton').grid(row=row, column=2, padx=(10,0), pady=4)

    def browse(self, var):
        path = filedialog.askopenfilename(title='Seleccionar listado CSV', filetypes=[('CSV', '*.csv'), ('Todos los archivos','*.*')])
        if path: var.set(path)

    def choose_drive_folder(self):
        folder = filedialog.askdirectory(title='Elegí la carpeta de Google Drive que contiene viejo.csv y nuevo.csv',
                                        initialdir=str(self.drive_folder) if self.drive_folder.is_dir() else None)
        if folder:
            self.drive_folder = Path(folder)
            save_drive_folder(self.drive_folder)
            messagebox.showinfo('Carpeta de Drive guardada', str(self.drive_folder))
        return bool(folder)

    def load_from_drive(self):
        try:
            if not self.drive_folder.is_dir():
                if not self.choose_drive_folder(): return
            old, new = find_drive_pair(self.drive_folder)
            self.old_path.set(str(old))
            self.new_path.set(str(new))
            self.compare()
        except Exception as exc:
            messagebox.showerror('No se pudieron cargar las listas de Drive', str(exc) +
                '\n\nComprobá que Drive haya terminado de sincronizar y que ambos archivos sean del mismo lote.')

    @staticmethod
    def _list_pair_id(old_path, new_path):
        dig = hashlib.sha256()
        for label, filename in [('viejo', old_path), ('nuevo', new_path)]:
            dig.update(label.encode('ascii'))
            with open(filename, 'rb') as src:
                for chunk in iter(lambda: src.read(1024 * 1024), b''):
                    dig.update(chunk)
        return dig.hexdigest()

    def _persist_print(self, codes, enabled):
        # Save both for backwards compatibility. The important record is by
        # article + final price, not CSV bytes: re-exporting identical prices
        # must NOT recheck the label the user rejected.
        self.db.set_print_decisions(self.current_lot, codes, enabled)
        catalog = {row['new']['code']: row['new'] for row in self.changes + self.new_items}
        updates = []
        for code in codes:
            product = catalog.get(code) or self.db.get_product(code)
            if product:
                updates.append((code, product['price'], enabled))
        if updates:
            self.db.set_print_preferences(updates)

    def compare(self):
        if not self.old_path.get() or not self.new_path.get():
            messagebox.showwarning('Faltan archivos', 'Seleccioná los dos listados.'); return
        try:
            self.configure(cursor='watch'); self.update_idletasks()
            before = load_csv(self.old_path.get())
            after = load_csv(self.new_path.get())
            lot = self._list_pair_id(self.old_path.get(), self.new_path.get())
            self.changes = [dict(new=entry, old=before[code]['price']) for code, entry in after.items() if code in before and entry['price'] != before[code]['price']]
            self.new_items = [dict(new=entry, old=None) for code, entry in after.items() if code not in before]
            self.db.sync_catalog(after)
            self.in_store = self.db.store_codes()
            self.changes.sort(key=lambda item: item['new']['name'].casefold())
            self.new_items.sort(key=lambda item: item['new']['name'].casefold())
            self.new_count_label.config(text=f'PRODUCTOS NUEVOS: {len(self.new_items):,}'.replace(',', '.'))
            # Cada comparación nueva selecciona por defecto para IMPRIMIR
            # únicamente los PRECIOS MODIFICADOS de productos marcados EN TIENDA.
            # Las dos selecciones son independientes: destildar impresión no
            # quita el artículo de tienda, ni se vuelve a tildar al refrescar.
            self.current_lot = lot
            self.chosen = {item['new']['code'] for item in self.changes
                           if item['new']['code'] in self.in_store}
            # Previous versions saved per exact file pair; preserve their
            # decisions for that pair and migrate them to the stronger rule.
            legacy = self.db.print_decisions(lot)
            for code, enabled in legacy.items():
                if enabled:
                    self.chosen.add(code)
                else:
                    self.chosen.discard(code)
            valid = {row['new']['code']: row['new']['price'] for row in self.changes + self.new_items}
            if legacy:
                self.db.set_print_preferences((code,valid[code], enabled)
                                              for code,enabled in legacy.items() if code in valid)
            # Strongest rule: decisions persist for a given product / price
            # even after re-exporting either CSV, using another PC's copies,
            # or re-running the comparison.
            for code, enabled in self.db.print_preferences(self.changes + self.new_items).items():
                if enabled:
                    self.chosen.add(code)
                else:
                    self.chosen.discard(code)
            self.qty.clear()
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
        path = filedialog.askopenfilename(title='Elegir catálogo inicial CSV o listado con stock',
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
        ttk.Label(container,text='Esta carga es PROVISIONAL. Las marcas existentes se conservan; las decisiones manuales tienen prioridad.',wraplength=665).pack(anchor='w',pady=(12,4))
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
                previous_store = set(self.in_store)
                if codes: self.db.set_store(codes, True, manual=False)
                self.in_store = self.db.store_codes()
                newly_in_store = self.in_store - previous_store
                already_rejected = {code for code, enabled in self.db.print_preferences(self.changes).items() if not enabled}
                already_rejected.update(code for code, enabled in self.db.print_decisions(self.current_lot).items() if not enabled)
                self.chosen.update(item['new']['code'] for item in self.changes
                                   if item['new']['code'] in newly_in_store and item['new']['code'] not in already_rejected)
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

    def import_barcode_csv(self):
        path = filedialog.askopenfilename(title='Importar códigos de barras e IdArticulo',
                filetypes=[('Listados CSV', '*.csv'), ('Todos los archivos', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8-sig', newline='') as f:
                sample = f.readline()
                f.seek(0)
                reader = csv.DictReader(f, delimiter=';' if sample.count(';') > sample.count(',') else ',')
                names = {str(k).strip().lower().replace(' ', '').replace('_', ''): k for k in reader.fieldnames or []}
                sku = next((names[n] for n in ('idarticulo', 'sku', 'codigointerno', 'codigo') if n in names), None)
                bar = next((names[n] for n in ('codigobarra', 'codigobarras', 'códigobarra', 'códigobarras', 'barcode', 'ean', 'ean13') if n in names), None)
                if not sku or not bar:
                    raise ValueError('El CSV debe tener las columnas IdArticulo y Codigo_Barra (o similares).')
                pairs = [(row.get(sku, ''), row.get(bar, '')) for row in reader]
            if not pairs:
                raise ValueError('El archivo no tiene registros.')
            if not messagebox.askyesno('Importar códigos de barras',
                    f'Se encontraron {len(pairs):,} filas.\n\nSe agregarán asociaciones código de barras ↔ código interno, sin reemplazar las existentes que correspondan a otro artículo. ¿Continuar?'):
                return
            count = self.db.import_barcodes(pairs)
            messagebox.showinfo('Códigos de barras importados',
                    f'Se incorporaron {count:,} asociaciones de código de barras. Ya podés escanear esos artículos.')
        except (OSError, ValueError, csv.Error, sqlite3.Error) as exc:
            messagebox.showerror('No se pudo importar el archivo', str(exc))

    def scan_product(self):
        """Search by internal code or description; USB keyboard-wedge scanner sends Enter."""
        dlg=tk.Toplevel(self)
        dlg.title('Cotyland · Buscar, escanear o ingresar artículo')
        dlg.geometry('830x540');dlg.minsize(670,410)
        dlg.transient(self)
        body=ttk.Frame(dlg,padding=14);body.pack(fill='both',expand=True)
        ttk.Label(body,text='Escaneá un código o escribí código / descripción y presioná ENTER',
                  font=('Segoe UI',11,'bold')).pack(anchor='w')
        query=tk.StringVar()
        entry=ttk.Entry(body,textvariable=query,font=('Segoe UI',13))
        entry.pack(fill='x',pady=9)
        cols=('code','name','price','store')
        tree=ttk.Treeview(body,columns=cols,show='headings',selectmode='browse')
        for key,title,width in [('code','Código',115),('name','Descripción',390),('price','Precio',130),('store','En tienda',80)]:
            tree.heading(key,text=title);tree.column(key,width=width,stretch=key=='name')
        tree.pack(fill='both',expand=True,pady=(0,8))
        hint=tk.StringVar(value='Al escanear, se buscan coincidencias en la base guardada.')
        ttk.Label(body,textvariable=hint).pack(anchor='w')
        def find(_=None):
            term=query.get().strip()
            tree.delete(*tree.get_children())
            matches=self.db.find_products(term)
            for code,name,price,store in matches:
                tree.insert('', 'end', iid=code, values=(code,name,display_money(money(price)), 'Sí' if store else 'No'))
            hint.set(f'{len(matches)} coincidencias. ' + ('Usá AGREGAR NUEVO si el código no figura.' if not matches else 'Elegí una fila para continuar.'))
            if len(matches)==1:tree.selection_set(matches[0][0]);tree.focus(matches[0][0])
        entry.bind('<Return>',find)
        def selected():
            keys=tree.selection()
            if not keys:
                messagebox.showinfo('Elegí un artículo','Buscá o escaneá y elegí el artículo de la lista.',parent=dlg)
                return None
            return keys[0]
        def link_unknown():
            barcode=query.get().strip()
            if not barcode:
                messagebox.showinfo('Escanear código', 'Escaneá primero el código de barras.', parent=dlg)
                return
            from tkinter.simpledialog import askstring
            internal=askstring('Vincular código de barras',
                'Código leído: ' + barcode + '\n\nIngresá el CÓDIGO INTERNO (IdArticulo / SKU) de este producto:',
                parent=dlg)
            if not internal:
                return
            try:
                self.db.link_barcode(barcode, internal)
                query.set(internal)
                find()
                hint.set('Código de barras ' + barcode + ' vinculado al artículo ' + internal + '. Se recordará en futuros escaneos.')
            except (ValueError, sqlite3.Error) as exc:
                messagebox.showerror('No se pudo vincular', str(exc), parent=dlg)

        def mark():
            code=selected()
            if not code:return
            self.save_store([code],True); self.refresh();find()
            hint.set(f'Artículo {code} marcado EN TIENDA y guardado permanentemente.')
        def print_one():
            code=selected()
            if not code:return
            self.chosen.add(code)
            self._persist_print([code], True)
            self.refresh()
            hint.set(f'Artículo {code} preparado para imprimir. Cerrá esta ventana y tocá CREAR ETIQUETAS PDF.')
        def edit():
            code=selected()
            if not code:return
            product=self.db.get_product(code)
            editor(product)
        def editor(product=None):
            win=tk.Toplevel(dlg);win.title('Agregar / editar artículo');win.geometry('580x320')
            win.transient(dlg);win.grab_set()
            area=ttk.Frame(win,padding=18);area.pack(fill='both',expand=True)
            vals=[tk.StringVar(value=product['code'] if product else query.get().strip()),
                  tk.StringVar(value=product['name'] if product else ''),
                  tk.StringVar(value=str(product['price']) if product else ''),
                  tk.StringVar(value='')]
            for ix,label in enumerate(['Código del artículo','Descripción','Precio de venta (ej. 1250,50)','Código de barras alternativo (opcional)']):
                ttk.Label(area,text=label).grid(row=ix,column=0,sticky='w',pady=9)
                ttk.Entry(area,textvariable=vals[ix],width=43).grid(row=ix,column=1,padx=10,sticky='ew',pady=9)
            def save():
                code,name,price,barcode=[v.get().strip() for v in vals]
                if not code or not name or not price:
                    messagebox.showerror('Faltan datos','Completá código, descripción y precio.',parent=win);return
                try:
                    self.db.save_product(code,name,price,True,barcode)
                    self.in_store=self.db.store_codes()
                    # The manual edit is available in the catalog immediately, not only after comparing.
                    self.refresh()
                    query.set(code);find();win.destroy()
                except Exception as exc:
                    messagebox.showerror('No se pudo guardar',str(exc),parent=win)
            ttk.Button(area,text='GUARDAR Y MARCAR EN TIENDA',command=save).grid(row=4,column=0,columnspan=2,pady=15)
            area.columnconfigure(1,weight=1)
        buttons=ttk.Frame(body);buttons.pack(fill='x',pady=(6,0))
        ttk.Button(buttons,text='BUSCAR',command=find).pack(side='left',padx=(0,7))
        ttk.Button(buttons,text='MARCAR EN TIENDA',command=mark).pack(side='left',padx=(0,7))
        ttk.Button(buttons,text='PREPARAR ETIQUETA',command=print_one).pack(side='left',padx=(0,7))
        ttk.Button(buttons,text='EDITAR PRECIO',command=edit).pack(side='left',padx=(0,7))
        ttk.Button(buttons,text='VINCULAR BARRAS A SKU',command=link_unknown).pack(side='right',padx=(6,0))
        ttk.Button(buttons,text='AGREGAR NUEVO',command=lambda:editor()).pack(side='right')
        entry.focus_set()

    def click_checkbox(self, event):
        row = self.table.identify_row(event.y)
        col = self.table.identify_column(event.x)
        if not row or self.table.identify_region(event.x, event.y) != 'cell': return
        if col == '#1':
            self.chosen.symmetric_difference_update({row})
            self._persist_print([row], row in self.chosen)
        elif col == '#2':
            adding = row not in self.in_store
            self.save_store([row], adding)
            if adding and row in {x['new']['code'] for x in self.changes + self.new_items}:
                self.chosen.add(row)
                self._persist_print([row], True)
        else: return
        self.refresh()

    def set_store(self, value):
        codes = set(self.table.selection())
        if not codes:
            messagebox.showinfo('Seleccioná filas', 'Seleccioná una o más filas con un clic (Ctrl o Shift para varias) y después usá este botón. También podés pulsar directamente la casilla En tienda.'); return
        self.save_store(codes, value)
        if value:
            newly_chosen = codes & {x['new']['code'] for x in self.changes + self.new_items}
            self.chosen.update(newly_chosen)
            self._persist_print(newly_chosen, True)
        self.refresh()

    def toggle_selected(self, event):
        row = self.table.identify_row(event.y)
        if not row or self.table.identify_column(event.x) in ('#1','#2'): return
        if row in self.chosen: self.chosen.remove(row)
        else: self.chosen.add(row)
        self._persist_print([row], row in self.chosen)
        self.refresh()

    def mark_visible(self):
        self.chosen.update(self.visible)
        self._persist_print(self.visible, True)
        self.refresh()

    def unmark(self):
        self._persist_print(self.chosen, False)
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
        for code in self.chosen:
            if code not in rows:
                p=self.db.get_product(code)
                if p:rows[code]={'new':p,'old':None}
        selected = []
        for code in self.chosen:
            if code not in rows: continue
            item = dict(rows[code]); item['new'] = dict(item['new'])
            manual_price = self.db.get_manual_price(code)
            if manual_price is not None: item['new']['price'] = manual_price
            selected.append(item)
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
