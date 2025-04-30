import os
import re
import csv
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox, VERTICAL, RIGHT, Y, LEFT, BOTH

# Regex baseado no padrão de nome do arquivo
REGEX = re.compile(
    r'^[^_]*_\d{2}_\d{8}_\d{11}_\d{2}-\d{2}-\d{4}\.pdf$'
)

def validar_arquivo(caminho):
    nome = os.path.basename(caminho)
    if REGEX.match(nome):
        return ('Válido', nome)
    else:
        return ('Inválido', caminho)

def validar_com_threads(pasta, output_path, tree, progressbar):
    arquivos_pdf = []
    for raiz, _, arquivos in os.walk(pasta):
        for nome in arquivos:
            if nome.lower().endswith('.pdf'):
                arquivos_pdf.append(os.path.join(raiz, nome))

    total = len(arquivos_pdf)
    if total == 0:
        messagebox.showinfo("Aviso", "Nenhum arquivo PDF encontrado.")
        return

    resultados_invalidos = []
    resultados_completos = []
    progresso = [0]

    def atualizar_barra():
        while progresso[0] < total:
            progressbar['value'] = (progresso[0] / total) * 100
            progressbar.update()
        progressbar['value'] = 100

    threading.Thread(target=atualizar_barra, daemon=True).start()

    with ThreadPoolExecutor(max_workers=20) as executor:
        futuros = {executor.submit(validar_arquivo, caminho): caminho for caminho in arquivos_pdf}
        for futuro in as_completed(futuros):
            resultado = futuro.result()
            if resultado:
                resultados_completos.append(resultado)
                if resultado[0] == 'Inválido':
                    resultados_invalidos.append(resultado[1])
            progresso[0] += 1

    # Atualiza a tabela após o processamento completo
    for status, caminho in resultados_completos:
        tree.insert('', 'end', values=(status, caminho))

    # Salva CSV
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['ARQUIVO'])
        for item in resultados_invalidos:
            writer.writerow([item])

    messagebox.showinfo("Concluído", f"Validação finalizada!\nArquivos inválidos salvos em:\n{output_path}")

def selecionar_pasta(entry_pasta):
    pasta = filedialog.askdirectory()
    if pasta:
        entry_pasta.delete(0, 'end')
        entry_pasta.insert(0, pasta)

def iniciar_busca(entry_pasta, tree, progressbar):
    pasta = entry_pasta.get()
    if not pasta or not os.path.isdir(pasta):
        messagebox.showerror("Erro", "Por favor, selecione uma pasta válida.")
        return

    tree.delete(*tree.get_children())
    progressbar['value'] = 0

    output_path = os.path.join(pasta, 'arquivos_invalidos.csv')
    threading.Thread(target=validar_com_threads, args=(pasta, output_path, tree, progressbar), daemon=True).start()

def main():
    app = ttk.Window(themename='darkly')
    app.title("Validador de PDFs - Nome de Arquivo")
    app.geometry("950x650")

    # Frame superior
    frame_top = ttk.Frame(app, padding=10)
    frame_top.pack(fill='x')

    ttk.Label(frame_top, text="Pasta de origem:").pack(side='left')
    entry_pasta = ttk.Entry(frame_top, width=70)
    entry_pasta.pack(side='left', padx=5, expand=True, fill='x')

    ttk.Button(frame_top, text="Selecionar Pasta", command=lambda: selecionar_pasta(entry_pasta)).pack(side='left', padx=5)
    ttk.Button(frame_top, text="Iniciar Busca", command=lambda: iniciar_busca(entry_pasta, tree, progressbar)).pack(side='left', padx=5)

    # Barra de progresso
    progressbar = ttk.Progressbar(app, mode='determinate')
    progressbar.pack(fill='x', padx=10, pady=(0, 10))

    # Frame da tabela
    frame_tree = ttk.Frame(app)
    frame_tree.pack(fill='both', expand=True, padx=10, pady=10)

    columns = ('Status', 'Arquivo')
    tree = ttk.Treeview(frame_tree, columns=columns, show='headings')
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor='w')

    # Scroll lateral
    scrollbar = ttk.Scrollbar(frame_tree, orient=VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree.pack(side=LEFT, fill=BOTH, expand=True)

    app.mainloop()

if __name__ == "__main__":
    main()
