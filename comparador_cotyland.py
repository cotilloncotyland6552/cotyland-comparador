import csv
import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime

APP_TITLE = 'Cotyland | Comparador de Precios y Etiquetas'
APP_VERSION = '7.0.0'


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
        self.stock_file = Path(os.environ.get('APPDATA', str(Path.home()))) / 'Cotyland_Comparador' / 'productos_en_tienda.json'
        self.in_store = self.load_store()
        self.only_store = tk.BooleanVar(value=False)
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
        cb = ttk.Combobox(bar, textvariable=self.kind, values=['Precios modificados','Artículos nuevos'], state='readonly', width=23)
        cb.pack(side='left', padx=(8,18)); cb.bind('<<ComboboxSelected>>', lambda _: self.refresh())
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
        footer = tk.Frame(self, bg='white', padx=18, pady=13)
        footer.pack(fill='x')
        ttk.Button(footer, text='Marcar visibles', command=self.mark_visible).pack(side='left', padx=(0,8))
        ttk.Button(footer, text='Desmarcar todas', command=self.unmark).pack(side='left', padx=(0,8))
        ttk.Button(footer, text='Marcar en tienda', command=lambda: self.set_store(True)).pack(side='left', padx=(0,8))
        ttk.Button(footer, text='Quitar de tienda', command=lambda: self.set_store(False)).pack(side='left', padx=(0,8))
        ttk.Button(footer, text='Copias seleccionadas…', command=self.set_copies).pack(side='left', padx=(0,8))
        ttk.Button(footer, text='Exportar cambios CSV', command=self.export).pack(side='right', padx=(8,0))
        ttk.Button(footer, text='CREAR ETIQUETAS PDF', command=self.labels).pack(side='right')

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
                    subprocess.Popen([sys.executable, str(Path(__file__).resolve())], cwd=str(Path(__file__).resolve().parent))
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
            self.changes.sort(key=lambda item: item['new']['name'].casefold())
            self.new_items.sort(key=lambda item: item['new']['name'].casefold())
            self.chosen.clear(); self.qty.clear()
            self.kind.set('Precios modificados'); self.search.set('')
            self.summary = f'{len(before):,} anteriores · {len(after):,} actuales · {len(self.changes):,} precios cambiados · {len(self.new_items):,} productos nuevos'.replace(',', '.')
            self.refresh()
        except Exception as e:
            messagebox.showerror('Error al leer los archivos', str(e))
        finally:
            self.configure(cursor='')

    def current(self):
        return self.changes if self.kind.get() == 'Precios modificados' else self.new_items

    def refresh(self):
        for item in self.table.get_children(): self.table.delete(item)
        q = self.search.get().strip().casefold()
        self.visible = {}
        for item in self.current():
            entry = item['new']; code = entry['code']
            if q and q not in code.casefold() and q not in entry['name'].casefold(): continue
            if self.only_store.get() and code not in self.in_store: continue
            self.visible[code] = item
            self.table.insert('', 'end', iid=code, values=('☑' if code in self.chosen else '□', '☑' if code in self.in_store else '□', code, entry['name'], display_money(item['old']) if item['old'] is not None else '—', display_money(entry['price']), self.qty.get(code,1)))
        self.status.set((getattr(self,'summary','Sin comparación') + f'   |   Mostrando {len(self.visible)}   |   Seleccionados {len(self.chosen)}'))

    def load_store(self):
        try:
            return set(json.loads(self.stock_file.read_text(encoding='utf-8')))
        except (OSError, ValueError, TypeError):
            return set()

    def save_store(self):
        try:
            self.stock_file.parent.mkdir(parents=True, exist_ok=True)
            self.stock_file.write_text(json.dumps(sorted(self.in_store), ensure_ascii=False), encoding='utf-8')
        except OSError as exc:
            messagebox.showwarning('No se pudo guardar', str(exc))

    def click_checkbox(self, event):
        row = self.table.identify_row(event.y)
        col = self.table.identify_column(event.x)
        if not row or self.table.identify_region(event.x, event.y) != 'cell': return
        if col == '#1':
            self.chosen.symmetric_difference_update({row})
        elif col == '#2':
            self.in_store.symmetric_difference_update({row})
            self.save_store()
        else: return
        self.refresh()

    def set_store(self, value):
        codes = set(self.table.selection())
        if not codes:
            messagebox.showinfo('Seleccioná filas', 'Seleccioná una o más filas con un clic (Ctrl o Shift para varias) y después usá este botón. También podés pulsar directamente la casilla En tienda.'); return
        if value: self.in_store.update(codes)
        else: self.in_store.difference_update(codes)
        self.save_store(); self.refresh()

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
            messagebox.showerror('Falta dependencia', 'Instalá ReportLab ejecutando INSTALAR_Y_ABRIR.bat.'); return
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

                # Precio con prioridad máxima. Objetivo 94 pt (doble que V5).
                # El alto de 35 mm impone un límite físico de 70 pt, y los
                # precios largos se reducen hasta entrar dentro de los 70 mm.
                # Los importes enteros se imprimen sin ,00 para ganar espacio.
                price = display_money(product['price'])
                target_size = 94.0
                vertical_limit = 70.0
                max_price_width = w - 1.5*mm
                size = min(target_size, vertical_limit)
                width_at_size = stringWidth(price, 'Helvetica-Bold', size)
                if width_at_size > max_price_width:
                    size *= max_price_width / width_at_size
                c.setFont('Helvetica-Bold', size)
                c.drawCentredString(x+w/2, y+10.5*mm, price)

                # Número de artículo legible, sin representación en barras.
                identifier = str(product['code']).strip()
                c.setFont('Helvetica-Bold', 8.2)
                c.drawCentredString(x+w/2, y+6.0*mm, clip('Cód. art.: ' + identifier, 'Helvetica-Bold', 8.2, w-4*mm))
                c.setFont('Helvetica', 7.0)
                c.drawCentredString(x+w/2, y+2.3*mm, 'Actualizado: ' + generated_date)
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
