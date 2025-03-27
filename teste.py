import os
import time  # Para medir o tempo de execução
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF para leitura rápida de PDFs
import threading
import concurrent.futures  # Para ProcessPoolExecutor
import ttkbootstrap as ttk  # Biblioteca para tema moderno
import re  # Para validação do termo de busca

# === TEMAS DISPONÍVEIS ===
temas_disponiveis = {
    "🔥 Noite Cyberpunk": "cyborg",  
    "🌞 Luz do Amanhã": "flatly",  
    "🌑 Sombras Profundas": "darkly",  
    "🦸 Herói Dourado": "superhero", 
    "☀️ Crepúsculo Solar": "solar",  
    "📜 Manuscrito Antigo": "journal",  
    "🏝️ Areias Douradas": "sandstone", 
    "🔮 Vibração Mística": "pulse",  
    "👨‍🚀 Horizonte Cósmico": "cosmo",  
    "⚡ Futurismo Elétrico": "morph"  
}
tema_atual = "🔥 Noite Cyberpunk"

def mudar_tema(event=None):
    """Altera o tema com base na seleção do combobox."""
    global tema_atual
    tema_atual = combobox_temas.get()
    janela.style.theme_use(temas_disponiveis[tema_atual])

def exibir_loading():
    """Cria uma janela independente de loading."""
    global janela_loading
    janela_loading = tk.Toplevel(janela)
    janela_loading.title("Processando...")
    janela_loading.geometry("300x150")
    janela_loading.resizable(False, False)
    ttk.Label(janela_loading, text="Processando PDFs...", font=("Arial", 12)).pack(pady=20)
    barra = ttk.Progressbar(janela_loading, mode="indeterminate")
    barra.pack(pady=10, padx=20, fill="x")
    barra.start(10)
    janela_loading.protocol("WM_DELETE_WINDOW", lambda: None)

def fechar_loading():
    """Fecha a janela de loading."""
    if janela_loading:
        janela_loading.destroy()

def validar_termo_busca(termo):
    """Verifica se o termo de busca está no formato correto (ex: 1823-35)."""
    return bool(re.fullmatch(r"\d{1,6}-\d{1,6}", termo))

def processar_pdf(pdf_path, termo_busca, pasta_saida):
    """Processa um único PDF e filtra páginas que contêm o termo de busca."""
    try:
        doc = fitz.open(pdf_path)
        pdf_writer = fitz.open()
        paginas_filtradas = []

        for num_pagina in range(len(doc)):
            texto = doc[num_pagina].get_text("text", flags=fitz.TEXTFLAGS_TEXT)
            if termo_busca in texto:
                pdf_writer.insert_pdf(doc, from_page=num_pagina, to_page=num_pagina)
                paginas_filtradas.append(num_pagina + 1)

        if paginas_filtradas:
            nome_pasta = os.path.basename(os.path.dirname(pdf_path))
            nome_base = os.path.splitext(os.path.basename(pdf_path))[0]
            nome_arquivo = f"{nome_pasta}_{nome_base}_filtrado.pdf"
            caminho_saida = os.path.join(pasta_saida, nome_arquivo)
            pdf_writer.save(caminho_saida)
            print(f"[✅] Salvo: {caminho_saida} (Páginas: {paginas_filtradas})")

        doc.close()
        pdf_writer.close()
    except Exception as e:
        print(f"[❌] Erro ao processar {pdf_path}: {e}")

def buscar_em_multiplos_pdfs():
    """Inicia a busca nos arquivos PDF selecionados e mede o tempo de execução."""
    termo = entrada_numero.get().strip()
    pasta_saida = entrada_saida.get().strip()
    arquivos_pdf = lista_pdfs.get(0, tk.END)

    if not termo or not pasta_saida:
        messagebox.showwarning("Aviso", "Preencha todos os campos!")
        return

    if not validar_termo_busca(termo):
        messagebox.showwarning("Erro", "O termo de busca deve estar no formato correto (ex: 1823-35).")
        return

    if not arquivos_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF selecionado.")
        return

    def executar():
        inicio = time.time()  # Inicia o temporizador

        with concurrent.futures.ProcessPoolExecutor() as executor:
            futures = [executor.submit(processar_pdf, pdf, termo, pasta_saida) for pdf in arquivos_pdf]
            concurrent.futures.wait(futures)

        tempo_total = time.time() - inicio  # Calcula o tempo total
        fechar_loading()
        messagebox.showinfo("Concluído", f"Todos os PDFs foram processados!\n⏳ Tempo total: {tempo_total:.2f} segundos")

    exibir_loading()
    thread = threading.Thread(target=executar)
    thread.start()

def selecionar_pdfs():
    """Abre a janela para seleção de múltiplos PDFs."""
    arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivos:
        for arquivo in arquivos:
            if arquivo not in lista_pdfs.get(0, tk.END):
                lista_pdfs.insert(tk.END, arquivo)

def remover_pdf():
    """Remove o arquivo PDF selecionado da lista."""
    try:
        selecionado = lista_pdfs.curselection()[0]
        lista_pdfs.delete(selecionado)
    except IndexError:
        messagebox.showwarning("Aviso", "Selecione um arquivo para remover.")

def limpar_lista():
    """Remove todos os arquivos da lista."""
    lista_pdfs.delete(0, tk.END)

def salvar_em():
    """Abre a janela para selecionar a pasta de saída dos PDFs processados."""
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_saida.delete(0, tk.END)
        entrada_saida.insert(0, pasta)

# === INTERFACE PRINCIPAL ===
janela = ttk.Window(themename=temas_disponiveis[tema_atual])
janela.title("🔍 Filtro de PDFs")
janela.geometry("700x600")
janela.minsize(600, 500)

frame_superior = ttk.Frame(janela)
frame_superior.pack(fill="x", pady=5)
ttk.Label(frame_superior, text="📄 Filtro de PDFs", font=("Arial", 14, "bold")).pack(side="left", padx=10)

# === SELEÇÃO DE TEMA ===
frame_tema = ttk.Frame(janela)
frame_tema.pack(pady=5)
ttk.Label(frame_tema, text="🎨 Escolha um tema:", font=("Arial", 12)).pack(side="left", padx=5)
combobox_temas = ttk.Combobox(frame_tema, values=list(temas_disponiveis.keys()), state="readonly")
combobox_temas.pack(side="left", padx=5)
combobox_temas.set(tema_atual)
combobox_temas.bind("<<ComboboxSelected>>", mudar_tema)

ttk.Label(janela, text="📂 Escolha múltiplos PDFs:", font=("Arial", 12)).pack(pady=5)

frame_lista = ttk.Frame(janela)
frame_lista.pack(pady=5, padx=10, fill="both", expand=True)

lista_pdfs = tk.Listbox(frame_lista, height=8, width=70, bg="#2b2b2b", fg="white", selectbackground="#007acc")
lista_pdfs.pack(fill="both", expand=True)

frame_botoes = ttk.Frame(janela)
frame_botoes.pack()
ttk.Button(frame_botoes, text="Selecionar PDFs", command=selecionar_pdfs, bootstyle="primary").pack(side="left", padx=5)
ttk.Button(frame_botoes, text="Remover Selecionado", command=remover_pdf, bootstyle="danger").pack(side="left", padx=5)
ttk.Button(frame_botoes, text="Limpar Lista", command=limpar_lista, bootstyle="warning").pack(side="left", padx=5)

# === ENTRADA DO TERMO DE BUSCA ===
ttk.Label(janela, text="🔎 Número exato para buscar:", font=("Arial", 12)).pack(pady=5)
entrada_numero = ttk.Entry(janela, width=40)
entrada_numero.pack(pady=5)

# === SELEÇÃO DA PASTA DE SAÍDA ===
ttk.Label(janela, text="💾 Pasta para salvar os PDFs:", font=("Arial", 12)).pack(pady=5)
frame_saida = ttk.Frame(janela)
frame_saida.pack(fill="x", padx=10)
entrada_saida = ttk.Entry(frame_saida, width=45)
entrada_saida.pack(side=tk.LEFT, fill="x", expand=True, padx=5)
ttk.Button(frame_saida, text="Selecionar", command=salvar_em, width=15, bootstyle="secondary").pack(side=tk.LEFT)

botao_buscar = ttk.Button(janela, text="📥 Buscar e Salvar PDFs 📤", command=buscar_em_multiplos_pdfs, bootstyle="success-outline")
botao_buscar.pack(pady=20)

if __name__ == "__main__":
    janela.mainloop()
