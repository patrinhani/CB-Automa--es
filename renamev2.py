import os
import re
from tkinter import Tk, filedialog, Button, Entry, Text, Scrollbar, messagebox
import tkinter as tk  # Importando o tk aqui para usar tk.DoubleVar()
from tkinter import ttk
import threading

# Função para verificar se o nome do arquivo está de acordo com o padrão
def verificar_nome_arquivo(nome_arquivo):
    erros = []
    nome_base, extensao = os.path.splitext(nome_arquivo)
    partes = nome_base.split('_')

    if len(partes) != 5:
        erros.append(f"Erro geral: O nome do arquivo '{nome_arquivo}' não segue o padrão esperado.")
        return erros, nome_arquivo  # Não tenta corrigir porque está mal formatado

    tipo, codigo, matricula, cpf, data = partes

    # Verificar tipo
    if tipo != "TERMO DE ÉTICA":
        erros.append(f"Erro no início: Deve começar com 'TERMO DE ÉTICA', mas veio '{tipo}'.")

    # Verificar código (deve conter 1 ou 2 dígitos numéricos)
    if not (codigo.isdigit() and len(codigo) in [1, 2]):
        erros.append(f"Erro na segunda parte: '{codigo}' deve conter 1 ou 2 dígitos.")
    
    # Verificar matrícula
    if not matricula.isdigit() or len(matricula) != 8:
        erros.append(f"Erro na matrícula: '{matricula}' deve ter exatamente 8 dígitos.")
        matricula = matricula.zfill(8)

    # Verificar CPF
    if not cpf.isdigit():
        erros.append(f"Erro no CPF: '{cpf}' contém caracteres inválidos.")
    elif len(cpf) < 11:
        erros.append(f"Erro no CPF: '{cpf}' foi ajustado para '{cpf.zfill(11)}'.")
        cpf = cpf.zfill(11)
    elif len(cpf) > 11:
        erros.append(f"Erro no CPF: '{cpf}' tem mais de 11 dígitos.")

    # Verificar data
    data_regex = r"^\d{2}-\d{2}-\d{4}$"
    if not re.match(data_regex, data):
        erros.append(f"Erro na data: '{data}' não está no formato dd-mm-aaaa.")
        data = "01-01-2025"  # Valor padrão

    # Reconstruir nome corrigido
    nome_corrigido = f"{tipo}_{codigo}_{matricula}_{cpf}_{data}{extensao}"

    return erros, nome_corrigido

# Função para verificar os arquivos na pasta
def verificar_arquivos(diretorio, progress_var):
    resultados = []
    arquivos_com_erro = 0  # Contador de arquivos com erro
    arquivos = os.listdir(diretorio)
    total_arquivos = len(arquivos)
    for i, arquivo in enumerate(arquivos):
        progress_var.set((i + 1) / total_arquivos * 100)  # Atualiza a barra de progresso
        root.update_idletasks()  # Permite atualizar a interface enquanto executa
        caminho_arquivo = os.path.join(diretorio, arquivo)
        if os.path.isfile(caminho_arquivo):
            erros, nome_correto = verificar_nome_arquivo(arquivo)
            if erros:
                arquivos_com_erro += 1  # Incrementa o contador de arquivos com erro
                resultados.append((arquivo, erros, nome_correto))
    progress_var.set(100)  # Barra de progresso 100% após a verificação
    return resultados, arquivos_com_erro

# Função para renomear os arquivos
def renomear_arquivos(diretorio, progress_var):
    arquivos_renomeados = 0
    arquivos = os.listdir(diretorio)
    total_arquivos = len(arquivos)
    for i, arquivo in enumerate(arquivos):
        progress_var.set((i + 1) / total_arquivos * 100)  # Atualiza a barra de progresso
        root.update_idletasks()  # Permite atualizar a interface enquanto executa

        caminho_arquivo = os.path.join(diretorio, arquivo)
        erros, nome_correto = verificar_nome_arquivo(arquivo)
        if erros:
            caminho_novo = os.path.join(diretorio, nome_correto)
            try:
                os.rename(caminho_arquivo, caminho_novo)
                arquivos_renomeados += 1
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao renomear '{arquivo}': {e}")
    progress_var.set(100)  # Barra de progresso 100% após a renomeação

    if arquivos_renomeados > 0:
        messagebox.showinfo("Sucesso", f"{arquivos_renomeados} arquivo(s) renomeado(s) com sucesso!")
    else:
        messagebox.showinfo("Sem mudanças", "Nenhum arquivo foi renomeado.")

# Função para rodar a verificação em uma thread separada
def verificar_thread():
    diretorio = entry_diretorio.get()
    if not diretorio:
        messagebox.showerror("Erro", "Por favor, selecione uma pasta com arquivos.")
        return

    progress_var.set(0)  # Zera a barra de progresso antes de começar
    resultados, arquivos_com_erro = verificar_arquivos(diretorio, progress_var)

    text_resultados.delete("1.0", "end")
    if resultados:
        text_resultados.insert("end", f"Arquivos com erro encontrados: {arquivos_com_erro}\n\n")
        for arquivo, erros, nome_correto in resultados:
            text_resultados.insert("end", f"Arquivo com erros: '{arquivo}':\n")
            for erro in erros:
                text_resultados.insert("end", f"    - {erro}\n")
            text_resultados.insert("end", f"Nome corrigido: '{nome_correto}'\n\n")
    else:
        text_resultados.insert("end", "Todos os arquivos estão no padrão esperado.")

# Função para rodar a renomeação em uma thread separada
def renomear_thread():
    diretorio = entry_diretorio.get()
    if not diretorio:
        messagebox.showerror("Erro", "Por favor, selecione uma pasta com arquivos.")
        return

    progress_var.set(0)  # Zera a barra de progresso antes de começar
    renomear_arquivos(diretorio, progress_var)

# Função para iniciar a verificação em thread
def iniciar_verificacao():
    thread = threading.Thread(target=verificar_thread)
    thread.start()

# Função para iniciar a renomeação em thread
def iniciar_renomeacao():
    thread = threading.Thread(target=renomear_thread)
    thread.start()

# Função para selecionar a pasta de arquivos
def selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    if diretorio:
        entry_diretorio.delete(0, "end")
        entry_diretorio.insert(0, diretorio)

# Configuração da interface gráfica
root = Tk()
root.title("Verificador de Arquivos")

# Layout da interface
label_diretorio = Button(root, text="Selecione a pasta com os arquivos", command=selecionar_diretorio)
label_diretorio.pack(padx=10, pady=5)

entry_diretorio = Entry(root, width=50)
entry_diretorio.pack(padx=10, pady=5)

button_verificar = Button(root, text="Verificar Arquivos", command=iniciar_verificacao)
button_verificar.pack(padx=10, pady=10)

button_renomear = Button(root, text="Renomear Arquivos", command=iniciar_renomeacao)
button_renomear.pack(padx=10, pady=10)

# Barra de progresso
progress_var = tk.DoubleVar()  # Agora usando tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400)
progress_bar.pack(padx=10, pady=10)

# Área de resultados
scrollbar = Scrollbar(root)
scrollbar.pack(side="right", fill="y")

text_resultados = Text(root, wrap="word", height=20, width=80, yscrollcommand=scrollbar.set)
text_resultados.pack(padx=10, pady=5, fill="both", expand=True)

scrollbar.config(command=text_resultados.yview)

# Tamanho inicial da janela
root.geometry("800x600")
root.resizable(True, True)

root.mainloop()
