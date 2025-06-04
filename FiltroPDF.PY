import os
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import threading
import queue
import ttkbootstrap as ttk
from collections import defaultdict

# === TEMAS DISPONÍVEIS ===
temas_disponíveis = {
    "🔥 Cyberpunk Noite": "cyborg",
    "🌞 Amanhecer Claro": "flatly",
    "🌑 Escuridão Profunda": "darkly",
    "🦸 Herói Moderno": "superhero",
    "☀️ Solar Elegante": "solar",
    "📜 Manuscrito Clássico": "journal",
    "🏝️ Praia Dourada": "sandstone",
    "🌊 Oceano Sereno": "united",
    "🔮 Misticismo Vibrante": "pulse",
    "👨‍🚀 Espaço Cósmico": "cosmo",
    "⚡ Futuro Elétrico": "morph"
}
tema_atual = "🔥 Cyberpunk Noite"

# === FUNÇÕES ===
def mudar_tema(event=None):
    global tema_atual
    tema_atual = combobox_temas.get()
    janela.style.theme_use(temas_disponíveis[tema_atual])

def exibir_loading():
    global janela_loading, barra_progresso
    janela_loading = tk.Toplevel(janela)
    janela_loading.title("Processando...")
    janela_loading.geometry("300x150")
    janela_loading.resizable(False, False)
    ttk.Label(janela_loading, text="Processando PDFs...", font=("Arial", 12)).pack(pady=20)
    barra_progresso = ttk.Progressbar(janela_loading, mode="determinate", maximum=100)
    barra_progresso.pack(pady=10, padx=20, fill="x")
    janela_loading.protocol("WM_DELETE_WINDOW", lambda: None)

def fechar_loading():
    if janela_loading:
        janela_loading.destroy()

def processar_pdfs(pdfs_por_pasta, termo_busca, pasta_saida, fila):
    total_pdfs = sum(len(arquivos) for arquivos in pdfs_por_pasta.values())
    processados = 0
    for pasta, arquivos in pdfs_por_pasta.items():
        try:
            pdf_writer = fitz.open()
            paginas_filtradas = []

            for pdf_path in arquivos:
                doc = fitz.open(pdf_path)
                for num_pagina in range(len(doc)):
                    texto = doc[num_pagina].get_text("text", flags=fitz.TEXTFLAGS_TEXT)
                    if termo_busca in texto:
                        pdf_writer.insert_pdf(doc, from_page=num_pagina, to_page=num_pagina)
                        paginas_filtradas.append(num_pagina + 1)
                doc.close()

                processados += 1
                progresso = int((processados / total_pdfs) * 100)
                fila.put(progresso)

            if paginas_filtradas:
                nome_pasta = os.path.basename(pasta)
                nome_arquivo = f"{nome_pasta}_filtrado.pdf"
                caminho_saida = os.path.join(pasta_saida, nome_arquivo)
                pdf_writer.save(caminho_saida)
                print(f"[✅] Salvo: {caminho_saida} (Páginas: {paginas_filtradas})")
            pdf_writer.close()
        except Exception as e:
            print(f"[❌] Erro ao processar PDFs em {pasta}: {e}")
    fila.put("done")

def buscar_em_multiplos_pdfs():
    termo = entrada_numero.get().strip()
    pasta_saida = entrada_saida.get().strip()
    arquivos_pdf = lista_pdfs.get(0, tk.END)

    if not termo or not pasta_saida:
        messagebox.showwarning("Aviso", "Preencha todos os campos!")
        return

    if not arquivos_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF selecionado.")
        return

    pdfs_por_pasta = defaultdict(list)
    for pdf in arquivos_pdf:
        pasta = os.path.dirname(pdf)
        pdfs_por_pasta[pasta].append(pdf)

    fila = queue.Queue()

    def mostrar_mensagem_temporaria():
        msg = tk.Toplevel(janela)
        msg.title("✅ Concluído")
        msg.geometry("300x100")
        msg.resizable(False, False)
        ttk.Label(msg, text="Todos os PDFs foram processados!", font=("Arial", 11)).pack(pady=20)
        msg.after(2000, msg.destroy)

    def executar():
        processar_pdfs(pdfs_por_pasta, termo, pasta_saida, fila)

    def verificar_fila():
        try:
            while True:
                valor = fila.get_nowait()
                if valor == "done":
                    fechar_loading()
                    mostrar_mensagem_temporaria()
                    return
                barra_progresso['value'] = valor
        except queue.Empty:
            janela.after(100, verificar_fila)

    exibir_loading()
    threading.Thread(target=executar, daemon=True).start()
    verificar_fila()

def selecionar_pdfs():
    arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivos:
        for arquivo in arquivos:
            if arquivo not in lista_pdfs.get(0, tk.END):
                lista_pdfs.insert(tk.END, arquivo)

def selecionar_pasta():
    pasta = filedialog.askdirectory()
    if pasta:
        for raiz, _, arquivos in os.walk(pasta):
            for nome in arquivos:
                if nome.endswith(".pdf"):
                    caminho = os.path.join(raiz, nome)
                    if caminho not in lista_pdfs.get(0, tk.END):
                        lista_pdfs.insert(tk.END, caminho)

def remover_pdf():
    try:
        selecionado = lista_pdfs.curselection()[0]
        lista_pdfs.delete(selecionado)
    except IndexError:
        messagebox.showwarning("Aviso", "Selecione um arquivo para remover.")

def salvar_em():
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_saida.delete(0, tk.END)
        entrada_saida.insert(0, pasta)

# === INTERFACE ===
janela = ttk.Window(themename=temas_disponíveis[tema_atual])
janela.title("🔍 Filtro de PDFs")
janela.geometry("700x600")
janela.minsize(600, 500)
janela.rowconfigure(2, weight=1)
janela.columnconfigure(0, weight=1)

frame_superior = ttk.Frame(janela)
frame_superior.pack(fill="x", pady=5)
ttk.Label(frame_superior, text="📄 Filtro de PDFs", font=("Arial", 14, "bold")).pack(side="left", padx=10)

frame_tema = ttk.Frame(janela)
frame_tema.pack(pady=5)
ttk.Label(frame_tema, text="🎨 Escolha um tema:", font=("Arial", 12)).pack(side="left", padx=5)
combobox_temas = ttk.Combobox(frame_tema, values=list(temas_disponíveis.keys()), state="readonly")
combobox_temas.pack(side="left", padx=5)
combobox_temas.set(tema_atual)
combobox_temas.bind("<<ComboboxSelected>>", mudar_tema)

frame_lista = ttk.Frame(janela)
frame_lista.pack(pady=5, padx=10, fill="both", expand=True)
scrollbar = ttk.Scrollbar(frame_lista, orient="vertical")
lista_pdfs = tk.Listbox(frame_lista, height=8, width=70)
lista_pdfs.pack(side=tk.LEFT, fill="both", expand=True)
scrollbar.pack(side=tk.RIGHT, fill="y")
lista_pdfs.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=lista_pdfs.yview)

frame_saida = ttk.Frame(janela)
frame_saida.pack(pady=5, padx=10, fill="x")
ttk.Label(frame_saida, text="📂 Pasta de saída:", font=("Arial", 12)).pack(side="left", padx=5)
entrada_saida = ttk.Entry(frame_saida, width=50)
entrada_saida.pack(side="left", padx=5)
ttk.Button(frame_saida, text="Salvar em...", command=salvar_em).pack(side="left", padx=5)

frame_termo = ttk.Frame(janela)
frame_termo.pack(pady=5, padx=10, fill="x")
ttk.Label(frame_termo, text="🔍 Termo de busca:", font=("Arial", 12)).pack(side="left", padx=5)
entrada_numero = ttk.Entry(frame_termo, width=50)
entrada_numero.pack(side="left", padx=5)

frame_botoes = ttk.Frame(janela)
frame_botoes.pack()
ttk.Button(frame_botoes, text="Selecionar PDFs", command=selecionar_pdfs).pack(side=tk.LEFT, padx=5)
ttk.Button(frame_botoes, text="Selecionar Pasta", command=selecionar_pasta).pack(side=tk.LEFT, padx=5)
ttk.Button(frame_botoes, text="Remover Selecionado", command=remover_pdf).pack(side=tk.LEFT, padx=5)

ttk.Button(janela, text="📥 Buscar e Salvar PDFs 📤", command=buscar_em_multiplos_pdfs).pack(pady=20)

if __name__ == "__main__":
    janela.mainloop()