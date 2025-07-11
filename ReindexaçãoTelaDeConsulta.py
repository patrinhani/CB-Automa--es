from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import tkinter as tk
from tkinter import scrolledtext
import time
import logging
from datetime import datetime
import os
import json

# =================================================================================
# BLOCO DE CONFIGURAÇÃO (sem alterações)
# =================================================================================
log_filename = f"automacao_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S", filename=log_filename, filemode="w")
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S"))
logging.getLogger().addHandler(console_handler)
logging.info(f"Todos os logs detalhados serão salvos em: {log_filename}")

download_directory = r"C:\Users\2160036544\Downloads\Automacao_Arquivos"
if not os.path.exists(download_directory):
    os.makedirs(download_directory)

# =================================================================================
# FUNÇÃO DE DOWNLOAD (sem alterações)
# =================================================================================
def realizar_download(page, context, matricula):
    popup_page = None
    try:
        results_frame = page.frame_locator("#frmViewer").frame_locator("#frameDocSelect")

        logging.info(f"[{matricula}] - Aguardando tabela de resultados...")
        first_row = results_frame.locator("table#primaryHitlist_grid tbody tr[role='row']").first
        try:
            first_row.wait_for(state="visible", timeout=20000)
            logging.info(f"[{matricula}] - Tabela de resultados carregada.")
        except PlaywrightTimeoutError:
            logging.warning(f"[{matricula}] - Nenhum documento encontrado na tabela para esta matrícula.")
            return True

        all_rows = results_frame.locator("table#primaryHitlist_grid tbody tr[role='row']").all()
        if not all_rows:
            logging.warning(f"[{matricula}] - Nenhum documento na tabela após espera. Pulando.")
            return True

        logging.info(f"[{matricula}] - Encontrados {len(all_rows)} documentos. Selecionando todos...")
        all_rows[0].click() 
        time.sleep(0.2)
        for i, row in enumerate(all_rows[1:]):
            row.click(modifiers=["Control"])
            time.sleep(0.1) 
        
        logging.info(f"[{matricula}] - Abrindo menu de contexto...")
        all_rows[-1].click(button="right")
        
        page.locator("li.contextMenuItem:has-text('Enviar para')").hover()
        time.sleep(0.5)
        
        logging.info(f"[{matricula}] - Aguardando a abertura do popup...")
        with context.expect_page() as popup_info:
            page.get_by_role("menuitem", name="Arquivo", exact=True).click()
            
        popup_page = popup_info.value
        popup_page.wait_for_load_state()
        logging.info(f"[{matricula}] - Popup capturado: '{popup_page.title()}'")
        
        format_selector = popup_page.locator("#selectContent")
        format_selector.wait_for(state="visible", timeout=30000)
        format_selector.select_option(value="pdf")
        time.sleep(1)
        
        logging.info(f"[{matricula}] - Aguardando o início do download...")
        with popup_page.expect_download(timeout=90000) as download_info:
            popup_page.locator("#btnSave").click()
            
        download = download_info.value
        suggested_filename = download.suggested_filename or "documentos.zip"
        new_filename = f"{matricula}_{suggested_filename}"
        download_path = os.path.join(download_directory, new_filename)
        download.save_as(download_path)
        
        logging.info(f"Download para a matrícula {matricula} concluído.")
        
        time.sleep(2)
        if os.path.exists(download_path) and os.path.getsize(download_path) > 0:
            logging.info(f"VERIFICAÇÃO CONCLUÍDA: O arquivo '{new_filename}' foi baixado com sucesso.")
        else:
            logging.error(f"FALHA NA VERIFICAÇÃO: O arquivo '{new_filename}' não foi encontrado ou está vazio.")
            raise Exception("Falha na verificação final do download.")
        
        return True

    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado durante o download para a matrícula {matricula}: {e}", exc_info=True)
        return False
    finally:
        if popup_page and not popup_page.is_closed():
            logging.info(f"[{matricula}] - Fechando a janela de popup...")
            popup_page.close()
            logging.info(f"[{matricula}] - Retornando para a página principal.")

# =================================================================================
# BLOCO DE EXECUÇÃO PRINCIPAL (COM LÓGICA DE LIMPEZA ATIVA)
# =================================================================================
def run_automation(playwright, matriculas_para_buscar):
    browser = None
    try:
        browser = playwright.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True, accept_downloads=True)
        page = context.new_page()

        logging.info("Acessando a página de login...")
        page.goto("https://im40333.onbaseonline.com/203appnet/Login.aspx")
        logging.info("Realizando login...")
        page.locator('xpath=/html/body/form/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/input').fill("60036544")
        page.locator('xpath=/html/body/form/div[3]/div/div[2]/table/tbody/tr[3]/td[2]/input').fill("12345678")
        page.locator('xpath=/html/body/form/div[3]/div/div[2]/table/tbody/tr[5]/td[2]/button').click()
        logging.info("Login realizado com sucesso.")

        # --- LOOP PRINCIPAL PARA PROCESSAR CADA MATRÍCULA ---
        for i, matricula in enumerate(matriculas_para_buscar):
            logging.info(f"================ INICIANDO PROCESSAMENTO PARA A MATRÍCULA: {matricula} ({i+1}/{len(matriculas_para_buscar)}) ================")
            
            try:
                nav_frame = page.frame_locator("#NavPanelIFrame")
                # Garante que estamos no menu correto antes de cada busca
                nav_frame.locator('xpath=/html/body/form/div[2]/div[1]/div/div/div/div[1]/div//input').fill("ANEXOS")
                time.sleep(1)
                nav_frame.locator('xpath=/html/body/form/div[2]/div[1]/div/div/div/div[3]/div[1]/ul/li[5]/label').click()
                
                # --- LÓGICA DE LIMPEZA ATIVA ---
                # A partir da segunda matrícula, limpa os resultados anteriores antes da nova busca
                if i > 0:
                    logging.info(f"[{matricula}] - Limpando resultados da busca anterior...")
                    results_frame = page.frame_locator("#frmViewer").frame_locator("#frameDocSelect")
                    tabela_body = results_frame.locator("table#primaryHitlist_grid tbody")
                    # Executa um script JS para apagar o conteúdo do corpo da tabela
                    if tabela_body.count() > 0:
                        tabela_body.evaluate("element => element.innerHTML = ''")
                        logging.info(f"[{matricula}] - Tabela de resultados limpa.")
                
                logging.info(f"[{matricula}] - Preenchendo campo de matrícula...")
                nav_frame.locator("div[keywordtypeid='114'] input.keywordInput").fill(matricula)
                
                logging.info(f"[{matricula}] - Clicando em pesquisar...")
                nav_frame.locator("button.js-searchButton:has-text('Pesquisar')").click()

                # Chama a função de download
                if not realizar_download(page, context, matricula):
                    logging.error(f"O processo de download falhou para a matrícula {matricula}. Pulando para a próxima.")
                
            except Exception as e:
                logging.error(f"Ocorreu uma falha geral no processamento da matrícula {matricula}: {e}", exc_info=True)
                logging.info("Tentando recarregar a página para continuar...")
                page.reload()
                time.sleep(5)
                
    except Exception as e:
        logging.error(f"Ocorreu um erro fatal na automação: {e}", exc_info=True)
    finally:
        if browser:
            browser.close()
        logging.info("Script finalizado.")

# --- PONTO DE ENTRADA DO SCRIPT (sem alterações) ---
if __name__ == "__main__":
    matriculas_finais = []
    def on_submit():
        raw_text = text_area.get("1.0", tk.END)
        processed_text = raw_text.replace(',', '\n')
        matriculas = [line.strip() for line in processed_text.split('\n') if line.strip()]
        if matriculas:
            global matriculas_finais
            matriculas_finais = matriculas
            root.destroy()
        else:
            root.destroy()

    root = tk.Tk()
    root.title("Entrada de Matrículas para Automação")
    label = tk.Label(root, text="Insira as matrículas (separadas por vírgula ou uma por linha):", padx=10, pady=10)
    label.pack()
    text_area = scrolledtext.ScrolledText(root, height=15, width=50)
    text_area.pack(padx=10, pady=5)
    text_area.focus()
    submit_button = tk.Button(root, text="Iniciar Automação", command=on_submit, height=2, width=20)
    submit_button.pack(padx=10, pady=10)
    root.eval('tk::PlaceWindow . center')
    root.mainloop()

    if matriculas_finais:
        logging.info(f"Matrículas a serem processadas: {matriculas_finais}")
        with sync_playwright() as playwright:
            run_automation(playwright, matriculas_finais)
    else:
        logging.warning("Nenhuma matrícula foi inserida. Encerrando o script.")