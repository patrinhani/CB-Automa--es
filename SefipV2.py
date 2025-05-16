import os
import re
import pandas as pd
import fitz  # PyMuPDF
from concurrent.futures import ProcessPoolExecutor
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from ttkbootstrap import Style, ttk
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
 
# Funções utilitárias
def limpar_cnpj(cnpj):
    return re.sub(r'\D', '', str(cnpj))
 
def extrair_comp(texto):
    match = re.search(r'COMP\s*[:\.\-]?\s*(\d{2})/(\d{4})', texto)
    if match:
        return match.group(1), match.group(2)
    return None, None
 
# Processamento de PDF
def processar_pdf(caminho_pdf, cnpj_original, cnpj_limpo):
    resultados = []
    if not os.path.exists(caminho_pdf):
        return resultados
    doc = fitz.open(caminho_pdf)
    for i in range(len(doc)):
        texto = doc.load_page(i).get_text() or ""
        if cnpj_original in texto or cnpj_limpo and cnpj_limpo in limpar_cnpj(texto):
            mes, ano = extrair_comp(texto)
            if mes and ano:
                resultados.append((ano, mes, caminho_pdf, i))
    doc.close()
    return resultados
 
# Função principal de processamento
def processar(caminho_excel, pasta_pdf_base, pasta_destino, status_text, progress_callback, pause_event, terminal):
    os.makedirs(pasta_destino, exist_ok=True)
    status_text.set("📚 Lendo lista de CNPJs...")
    terminal.insert(tk.END, f"Lendo Excel: {caminho_excel}\n")
 
    # Carrega e formata DataFrame
    df = pd.read_excel(caminho_excel)
    total_linhas = len(df)
    anos_para_buscar = [d for d in os.listdir(pasta_pdf_base)
                        if os.path.isdir(os.path.join(pasta_pdf_base, d))]
 
    # Prepara planilha formatada: salva e aplica estilos no header
    df.to_excel(caminho_excel, index=False)
    wb = load_workbook(caminho_excel)
    ws = wb.active
    header_font = Font(bold=True, color="FFFFFF")
    header_align = Alignment(horizontal='center')
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_align
        cell.fill = header_fill
    wb.save(caminho_excel)
    wb.close()
 
    # Calcula passos totais para progresso
    total_steps = sum(
        total_linhas for ano in anos_para_buscar
        for mes in range(1, 13)
        if os.path.isdir(os.path.join(pasta_pdf_base, ano, f"{mes:02d}_{ano}"))
    )
    progress_callback('configure', maximum=total_steps)
    step = 0
    start_time = time.time()
 
    # Loop por ano e mês
    for ano_loop in anos_para_buscar:
        for mes_num in range(1, 13):
            progress_callback('new_month', total_month=total_linhas)
            mes_str = f"{mes_num:02d}"
            status_text.set(f"🚀 Processando {mes_str}/{ano_loop}")
            pasta_mes = os.path.join(pasta_pdf_base, ano_loop, f"{mes_str}_{ano_loop}")
            if not os.path.isdir(pasta_mes):
                continue
            pdfs = [os.path.join(pasta_mes, f)
                    for f in os.listdir(pasta_mes)
                    if f.lower().endswith('.pdf')]
            terminal.insert(tk.END, f"Mês {mes_str}/{ano_loop}: {len(pdfs)} PDFs encontrados\n")
 
            for idx, row in df.iterrows():
                # Aguarda se pausado
                while pause_event.is_set():
                    time.sleep(0.1)
 
                col = f"{mes_str}_{ano_loop}"
                if str(row.get(col, '')).strip().lower() == 'concluído':
                    step += 1
                    progress_callback('step', value=1, start=start_time)
                    continue
 
                cnpj = str(row['cnpj'])
                cnpj_lim = limpar_cnpj(cnpj)
                resultados = {}
 
                # Processa PDFs em paralelo
                with ProcessPoolExecutor() as executor:
                    futures = [executor.submit(processar_pdf, p, cnpj, cnpj_lim) for p in pdfs]
                    for fut in futures:
                        for ano, mes, path, pag in fut.result():
                            resultados.setdefault((ano, mes), []).append((path, pag))
 
                # Se encontrou, gera novo PDF
                if resultados:
                    pasta_filial = os.path.join(pasta_destino, f"Filial - {row['Filial']}")
                    os.makedirs(pasta_filial, exist_ok=True)
                    for (ano, mes), lst in resultados.items():
                        novo_pdf = fitz.open()
                        textos_vistos = set()
                        for path, pag in lst:
                            texto = fitz.open(path).load_page(pag).get_text()
                            h = hash(texto)
                            if h in textos_vistos:
                                continue
                            textos_vistos.add(h)
                            novo_pdf.insert_pdf(fitz.open(path), from_page=pag, to_page=pag)
                        out = os.path.join(pasta_filial, ano, f"SEFIP - {mes}.pdf")
                        os.makedirs(os.path.dirname(out), exist_ok=True)
                        novo_pdf.save(out)
                        novo_pdf.close()
 
                        # Atualiza status na planilha
                        df.at[idx, col] = 'Concluído'
                        df.to_excel(caminho_excel, index=False)
 
                        # Log no terminal
                        filename = os.path.basename(out)
                        filial = row['Filial']
                        terminal.insert(tk.END, f"Filial {filial} - Ano {ano} - Arquivo {filename}\n")
                else:
                    df.at[idx, col] = 'Não encontrado'
                    df.to_excel(caminho_excel, index=False)
 
                # Atualiza progresso
                step += 1
                progress_callback('step', value=1, start=start_time)
            terminal.see(tk.END)
 
    status_text.set("🎉 Processo finalizado com sucesso!")
    terminal.insert(tk.END, "Processo completo.\n")
    terminal.see(tk.END)
 
# Interface gráfica com ttkbootstrap e terminal de logs
def iniciar_interface():
    style = Style(theme='vapor')
    root = style.master
    root.title("SEFIP")
    root.geometry("800x600")
 
    # Variáveis de controle
    vars_vars = {n: tk.StringVar(master=root) for n in ['excel', 'base', 'dest']}
    status = tk.StringVar(master=root, value="Aguardando início...")
    pause_event = threading.Event()
    total = tk.IntVar(master=root)
    month_total = tk.IntVar(master=root)
    month_current = tk.IntVar(master=root)
 
    # Frame de seleção de caminhos
    top = ttk.Frame(root, padding=10)
    top.pack(fill='x')
    campos = [('Excel', 'excel', [('Excel','*.xlsx')]), ('Pasta Base', 'base', None), ('Pasta Destino', 'dest', None)]
    for i, (label, key, ftypes) in enumerate(campos):
        ttk.Label(top, text=f"{label}:").grid(row=i, column=0, sticky='w')
        ttk.Entry(top, textvariable=vars_vars[key], width=60).grid(row=i, column=1)
        def cmd(v=key, ft=ftypes):
            if v == 'excel':
                path = filedialog.askopenfilename(filetypes=ft)
            else:
                path = filedialog.askdirectory()
            vars_vars[v].set(path)
        ttk.Button(top, text="Selecionar", command=cmd).grid(row=i, column=2)
 
    # Barra de progresso e indicadores
    bar = ttk.Progressbar(root, orient='horizontal', length=750, mode='determinate')
    bar.pack(pady=5)
    info_frame = ttk.Frame(root)
    info_frame.pack(fill='x')
    pct = ttk.Label(info_frame, text="0%")
    pct.pack(side='left')
    eta_lbl = ttk.Label(info_frame, text="ETA: --:--:--")
    eta_lbl.pack(side='left', padx=10)
    linha_lbl = ttk.Label(info_frame, text="0/0")
    linha_lbl.pack(side='left')
 
    # Terminal de logs
    log_frame = ttk.Frame(root)
    log_frame.pack(fill='both', expand=True)
    terminal = scrolledtext.ScrolledText(log_frame, height=15)
    terminal.pack(fill='both', expand=True, padx=5, pady=5)
 
    # Função de callback de progresso
    def progress_callback(action, **kwargs):
        if action == 'configure':
            maxv = kwargs.get('maximum', 1)
            bar.config(maximum=maxv, value=0)
            total.set(maxv)
        elif action == 'new_month':
            mt = kwargs.get('total_month', 0)
            month_total.set(mt)
            month_current.set(0)
            linha_lbl.config(text=f"0/{mt}")
        elif action == 'step':
            inc = kwargs.get('value', 1)
            bar.step(inc)
            v = bar['value']
            month_current.set(month_current.get() + inc)
            linha_lbl.config(text=f"{month_current.get()}/{month_total.get()}")
            pct.config(text=f"{v/total.get()*100:.1f}%")
            elapsed = time.time() - kwargs.get('start', time.time())
            rem = (elapsed / v * (total.get() - v)) if v else 0
            eta_lbl.config(text="ETA: " + time.strftime('%H:%M:%S', time.gmtime(rem)))
 
    # Botões iniciar e pausar
    def toggle_pause():
        if pause_event.is_set():
            pause_event.clear()
            btn_pause.config(text="Pausar")
        else:
            pause_event.set()
            btn_pause.config(text="Retomar")
 
    btn_pause = ttk.Button(root, text="Pausar", command=toggle_pause)
    btn_pause.pack(side='left', padx=5)
 
    def start():
        if not all(vars_vars.values()):
            messagebox.showerror("Erro", "Selecione todos os caminhos antes de iniciar.")
            return
        btn_start.config(state='disabled')
        threading.Thread(target=lambda: processar(
            vars_vars['excel'].get(), vars_vars['base'].get(), vars_vars['dest'].get(),
            status, progress_callback, pause_event, terminal
        ), daemon=True).start()
 
    btn_start = ttk.Button(root, text="Iniciar", command=start)
    btn_start.pack(side='left', padx=5)
    ttk.Label(root, textvariable=status).pack(side='right', padx=10)
 
    root.mainloop()
 
if __name__ == '__main__':
    iniciar_interface()
 