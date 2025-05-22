import os
import threading
import zipfile
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, BooleanVar, Canvas, Scrollbar, Label

class CompressorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Sefip")
        tb.Style(theme="vapor")

        # inicializações para evitar atributos faltantes
        self.source_dir = None
        self.target_dir = None
        self.filial_vars = {}
        self.year_vars = {}
        self._filtered_filiais = []
        self._years_data = {}

        # Top controls
        top = tb.Frame(master, padding=10)
        top.pack(fill='x')
        tb.Button(top, text="Selecionar Pasta Principal", bootstyle=SUCCESS,
                  command=self.select_source).pack(side='left')
        self.src_label = Label(top, text="Nenhuma pasta selecionada")
        self.src_label.pack(side='left', padx=10)
        tb.Button(top, text="Selecionar Pasta de Destino", bootstyle=INFO,
                  command=self.select_target).pack(side='left', padx=(20,0))
        self.dst_label = Label(top, text="Nenhuma pasta de destino selecionada")
        self.dst_label.pack(side='left', padx=10)

        # Main panels
        main = tb.Frame(master)
        main.pack(fill='both', expand=True, padx=10, pady=10)

        # Filiais panel
        filial_panel = tb.Frame(main)
        filial_panel.pack(side='left', fill='both', expand=True, padx=(0,5))
        Label(filial_panel, text="Filtrar Filial (número):").pack(anchor='w')
        self.filial_filter = tb.Entry(filial_panel)
        self.filial_filter.pack(fill='x', pady=(0,5))
        self.filial_filter.bind('<KeyRelease>', lambda e: self.filter_filiais())
        Label(filial_panel, text="Filiais Disponíveis:").pack(anchor='w')
        self.filial_canvas = Canvas(filial_panel)
        self.filial_canvas.pack(side='left', fill='both', expand=True)
        filial_scroll = Scrollbar(filial_panel, orient='vertical', command=self.filial_canvas.yview)
        filial_scroll.pack(side='right', fill='y')
        self.filial_canvas.configure(yscrollcommand=filial_scroll.set)
        self.filial_inner = tb.Frame(self.filial_canvas)
        self.filial_canvas.create_window((0,0), window=self.filial_inner, anchor='nw')
        self.filial_inner.bind('<Configure>', lambda e: self.filial_canvas.configure(scrollregion=self.filial_canvas.bbox('all')))
        self.filial_canvas.bind('<Configure>', lambda e: self.layout_filiais())

        # Years panel
        year_panel = tb.Frame(main)
        year_panel.pack(side='right', fill='both', expand=True, padx=(5,0))
        Label(year_panel, text="Anos por Filial:").pack(anchor='w')
        self.year_canvas = Canvas(year_panel)
        self.year_canvas.pack(side='left', fill='both', expand=True)
        year_scroll = Scrollbar(year_panel, orient='vertical', command=self.year_canvas.yview)
        year_scroll.pack(side='right', fill='y')
        self.year_canvas.configure(yscrollcommand=year_scroll.set)
        self.year_inner = tb.Frame(self.year_canvas)
        self.year_canvas.create_window((0,0), window=self.year_inner, anchor='nw')
        self.year_inner.bind('<Configure>', lambda e: self.year_canvas.configure(scrollregion=self.year_canvas.bbox('all')))
        self.year_canvas.bind('<Configure>', lambda e: self.layout_years())

        # Bottom controls
        bottom = tb.Frame(master, padding=10)
        bottom.pack(fill='x')
        self.progress = tb.Progressbar(bottom, orient=HORIZONTAL, mode="determinate")
        self.progress.pack(fill='x', pady=(0,5))
        tb.Button(bottom, text="Compactar Selecionados", bootstyle=PRIMARY,
                  command=self.start_compress).pack(fill='x')

    def select_source(self):
        path = filedialog.askdirectory(title="Selecione a pasta principal")
        if not path: return
        self.source_dir = path
        self.src_label.config(text=path)
        self.load_filiais()
        self.clear_years()

    def select_target(self):
        path = filedialog.askdirectory(title="Selecione a pasta de destino")
        if not path: return
        self.target_dir = path
        self.dst_label.config(text=path)

    def load_filiais(self):
        self.filial_vars = {}
        self._all_filiais = [d for d in sorted(os.listdir(self.source_dir))
                              if os.path.isdir(os.path.join(self.source_dir, d))]
        self.filter_filiais()

    def filter_filiais(self):
        for w in self.filial_inner.winfo_children(): w.destroy()
        txt = self.filial_filter.get().strip().lower()
        self._filtered_filiais = [d for d in self._all_filiais if not txt or txt in d.lower()]
        self.layout_filiais()

    def layout_filiais(self):
        # evita erro antes de ter dados
        if not hasattr(self, '_filtered_filiais'):
            return
        for w in self.filial_inner.winfo_children(): w.destroy()
        width = self.filial_canvas.winfo_width() or 300
        cols = max(1, width // 120)
        for i, d in enumerate(self._filtered_filiais):
            var = self.filial_vars.get(d, BooleanVar())
            cb = tb.Checkbutton(self.filial_inner, text=d, variable=var,
                                 command=self.on_filial_toggle)
            r, c = divmod(i, cols)
            cb.grid(row=r, column=c, sticky='w', padx=5, pady=2)
            self.filial_vars[d] = var
        self.filial_inner.update_idletasks()

    def clear_years(self):
        self.year_vars = {}
        for w in self.year_inner.winfo_children(): w.destroy()

    def load_years(self):
        self._years_data = {
            f: [y for y in sorted(os.listdir(os.path.join(self.source_dir, f)))
                if os.path.isdir(os.path.join(self.source_dir, f, y))]
            for f, var in self.filial_vars.items() if var.get()
        }
        self.layout_years()

    def layout_years(self):
        if not hasattr(self, '_years_data'):
            return
        for w in self.year_inner.winfo_children(): w.destroy()
        width = self.year_canvas.winfo_width() or 300
        cols = max(1, width // 120)
        for fi, (filial, years) in enumerate(self._years_data.items()):
            frame = tb.Labelframe(self.year_inner, text=filial, padding=5)
            frame.grid(row=fi, column=0, sticky='we', pady=5, padx=5)
            for j, y in enumerate(years):
                var = self.year_vars.get(filial, {}).get(y, BooleanVar())
                cb = tb.Checkbutton(frame, text=y, variable=var)
                r, c = divmod(j, cols)
                cb.grid(row=r, column=c, sticky='w', padx=3, pady=2)
                self.year_vars.setdefault(filial, {})[y] = var
        self.year_inner.update_idletasks()

    def on_filial_toggle(self):
        self.load_years()

    def start_compress(self):
        if not self.source_dir or not self.target_dir:
            messagebox.showwarning('Aviso', 'Selecione pasta principal e destino.')
            return
        combos = [(f, y) for f, ys in self.year_vars.items() for y, v in ys.items() if v.get()]
        if not combos:
            messagebox.showwarning('Aviso', 'Selecione pelo menos um ano em alguma filial.')
            return
        threading.Thread(target=self.compress, args=(combos,), daemon=True).start()

    def compress(self, combos):
        grouped = {}
        for filial, ano in combos:
            base = os.path.join(self.source_dir, filial, ano)
            for root, _, fs in os.walk(base):
                for f in fs:
                    grouped.setdefault(filial, []).append((root, f))
        total = sum(len(v) for v in grouped.values())
        self.progress.config(maximum=total, value=0)
        try:
            for filial, files in grouped.items():
                zip_path = os.path.join(self.target_dir, f"{filial}.zip")
                with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                    for _, (root, f) in enumerate(files, 1):
                        full = os.path.join(root, f)
                        arc = os.path.relpath(full, self.source_dir)
                        zf.write(full, arc)
                        self.progress.step()
        except Exception as e:
            messagebox.showerror('Erro', str(e))
        else:
            messagebox.showinfo('Sucesso', f'Zips salvos em {self.target_dir}')
        finally:
            self.progress.config(value=0)

if __name__ == '__main__':
    root = tb.Window()
    CompressorApp(root)
    root.mainloop()
