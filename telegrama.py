import os
import threading
import queue
import tempfile
import shutil
import re
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import extract_msg
from pathlib import Path

class MsgExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📨 Extrator Recursivo de Anexos de .MSG")
        self.root.geometry("800x600")
        self.root.columnconfigure(0, weight=1)

        self.style = ttk.Style(theme="darkly")

        self.input_path = ttk.StringVar()
        self.output_path = ttk.StringVar()

        frame_sel = ttk.Frame(self.root)
        frame_sel.pack(fill="x", pady=10, padx=10)

        ttk.Label(frame_sel, text="📂 Pasta de .MSG:", font=("Arial", 11)).grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_sel, textvariable=self.input_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(frame_sel, text="Selecionar...", command=self.selecionar_input, bootstyle=PRIMARY).grid(row=0, column=2)

        ttk.Label(frame_sel, text="📂 Pasta de saída:", font=("Arial", 11)).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame_sel, textvariable=self.output_path, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame_sel, text="Selecionar...", command=self.selecionar_output, bootstyle=PRIMARY).grid(row=1, column=2)

        ttk.Button(self.root, text="📥 Extrair Anexos", command=self.iniciar_extracao, bootstyle=SUCCESS).pack(pady=5)

        self.prog = ttk.Progressbar(self.root, mode="determinate", maximum=100, bootstyle=INFO)
        self.prog.pack(fill="x", padx=20)

        self.log_box = ScrolledText(self.root, height=20, font=("Consolas", 10), bg="#1e1e1e", fg="#ffffff", insertbackground="#ffffff")
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

        self.queue = queue.Queue()

    def selecionar_input(self):
        path = filedialog.askdirectory()
        if path:
            self.input_path.set(path)

    def selecionar_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path.set(path)

    def iniciar_extracao(self):
        inp = self.input_path.get().strip()
        out = self.output_path.get().strip()
        if not inp or not out:
            messagebox.showwarning("Aviso", "Selecione pastas de origem e destino.")
            return
        self.log_box.delete("1.0", "end")
        self.prog['value'] = 0
        threading.Thread(target=self.extrair_msgs, args=(inp, out), daemon=True).start()
        self.root.after(100, self.processar_queue)

    def extrair_msgs(self, inp, out):
        os.makedirs(out, exist_ok=True)
        arquivos = [f for f in os.listdir(inp) if f.lower().endswith('.msg')]
        total = len(arquivos)
        extraidos = [0]
        ignorados = []

        tempdir = Path(tempfile.mkdtemp())

        def sanitize(name: str) -> str:
            name = re.sub(r'[<>:"/\\|?*]+', '_', name)
            return re.sub(r'[^\w\-\.]+', '_', name)[:200]

        def save_attachment(att, target_dir: Path, filename: str) -> str:
            filename = sanitize(os.path.basename(filename))
            destino = target_dir / filename
            stem, ext = destino.stem, destino.suffix
            cnt = 1
            while destino.exists():
                destino = target_dir / f"{stem}_{cnt}{ext}"
                cnt += 1
            try:
                att.save(str(destino))
            except (TypeError, AttributeError):
                try:
                    # Se for bytes, salva diretamente
                    if hasattr(att, "data") and isinstance(att.data, bytes):
                        with open(destino, "wb") as f:
                            f.write(att.data)
                    # Se for um extract_msg.Message, salva como bytes
                    elif isinstance(att.data, extract_msg.Message):
                        with open(destino, "wb") as f:
                            f.write(att.data.as_bytes())
                    else:
                        raise ValueError("Formato de dados não suportado.")
                except Exception as e:
                    self.queue.put(f"[❌] Falha ao salvar {filename}: {e}\n")
                    return f"Erro ao salvar: {filename}"
            return destino.name

        def process_msg(msg_path: Path, top_name: str):
            try:
                m = extract_msg.Message(str(msg_path))
                for idx, att in enumerate(m.attachments, 1):
                    orig = getattr(att, 'longFilename', None) or getattr(att, 'filename', None) or getattr(att, 'shortFilename', None) or f"anexo_{idx}"
                    ext = Path(orig).suffix.lower()
                    if ext == '.msg':
                        nested_name = f"{top_name}_anexo{idx}.msg"
                        nested_path = tempdir / nested_name
                        saved_name = save_attachment(att, tempdir, nested_name)
                        self.queue.put(f"[→] .msg aninhado: {saved_name}\n")
                        process_msg(tempdir / saved_name, f"{top_name}_anexo{idx}")
                    else:
                        fname = f"{top_name}_anexo{idx}{ext}"
                        saved = save_attachment(att, Path(out), fname)
                        extraidos[0] += 1
                        self.queue.put(f"[✓] Salvo: {saved}\n")
            except Exception as e:
                self.queue.put(f"[❌] Erro em {msg_path.name}: {e}\n")

        for i, nome in enumerate(arquivos, 1):
            path_msg = Path(inp) / nome
            if not path_msg.exists():
                ignorados.append(nome)
                self.queue.put(f"[!] Não encontrado: {nome}\n")
            else:
                base = sanitize(path_msg.stem)
                process_msg(path_msg, base)
            self.queue.put(('PROG', int(i / total * 100)))

        shutil.rmtree(tempdir, ignore_errors=True)

        if ignorados:
            with open(Path(out) / 'ignorados.txt', 'w', encoding='utf-8') as f:
                f.writelines(x + '\n' for x in ignorados)
            self.queue.put("Ignorados registrados em: ignorados.txt\n")

        self.queue.put(f"\nTotal de .msg: {total}\n")
        self.queue.put(f"Anexos extraídos: {extraidos[0]}\n")
        self.queue.put(('DONE', None))

    def processar_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                if isinstance(item, tuple):
                    if item[0] == 'PROG':
                        self.prog['value'] = item[1]
                    elif item[0] == 'DONE':
                        messagebox.showinfo("Concluído", "Extração finalizada!")
                else:
                    self.log_box.insert("end", item)
                    self.log_box.see("end")
        except queue.Empty:
            pass
        finally:
            if self.prog['value'] < 100:
                self.root.after(100, self.processar_queue)

if __name__ == "__main__":
    app = MsgExtractorApp(ttk.Window(themename="darkly"))
    app.root.mainloop()
