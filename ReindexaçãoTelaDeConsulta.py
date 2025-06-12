import requests
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import keyboard
import time
import logging
from datetime import datetime
import os
import json

# ... (Configuração do Log, Driver e Funções de Falha permanecem iguais) ...
# ========== CONFIGURAÇÃO DO LOG (.txt) ==========
log_filename = f"automacao_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S", filename=log_filename, filemode="w")
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S"))
logging.getLogger().addHandler(console_handler)
logging.info(f"Todos os logs detalhados serão salvos em: {log_filename}")

# ========== CONFIGURAÇÃO DO EDGE DRIVER ==========
driver_path = r"C:\Users\2160036544\Downloads\edgedriver_win64\msedgedriver.exe"
options = Options()

# ============ CONFIGURAÇÕES DE DOWNLOAD ============
download_directory = r"C:\Users\2160036544\Downloads\Automacao_Arquivos"
if not os.path.exists(download_directory):
    os.makedirs(download_directory)
prefs = {
    "download.default_directory": download_directory,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.popups": 1
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-popup-blocking")

logging.info("Inicializando o Edge WebDriver...")
service = Service(driver_path)
driver = webdriver.Edge(service=service, options=options)
wait = WebDriverWait(driver, 40)

# ========== LÓGICA DE TRATAMENTO DE FALHAS ==========
def salvar_conteudo_para_depuracao(filename_prefix, content):
    try:
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_{current_time}.html"
        filepath = os.path.join(os.getcwd(), filename)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        logging.info(f"HTML/Conteúdo da falha salvo em: {filepath}")
    except Exception as e:
        logging.error(f"FALHA CRÍTICA ao salvar o HTML de depuração: {e}", exc_info=True)

def tratar_falha(driver_instance, contexto_da_falha, exception_obj, content_to_save=""):
    mensagem_erro = f"FALHA no contexto '{contexto_da_falha}': {exception_obj}"
    logging.error(mensagem_erro, exc_info=True)
    nome_arquivo_falha = f"FALHA_{contexto_da_falha.replace(' ', '_').replace('-', '_')}"
    if not content_to_save and driver_instance:
        try:
            content_to_save = driver_instance.page_source
        except:
            content_to_save = "Não foi possível obter o page_source."
    salvar_conteudo_para_depuracao(nome_arquivo_falha, content_to_save)

# ========== FUNÇÕES DE AUTOMAÇÃO ==========
def preencher_campo_rh(driver_instance, keyword_type_id, value_to_send, field_name):
    try:
        logging.info(f"Tentando preencher o campo '{field_name}'...")
        parent_div = wait.until(EC.visibility_of_element_located((By.XPATH, f"//div[@keywordtypeid='{keyword_type_id}']")))
        campo_input = wait.until(lambda d: parent_div.find_element(By.XPATH, ".//input[contains(@class, 'keywordInput')]"))
        campo_input.clear()
        campo_input.send_keys(value_to_send)
        logging.info(f"Campo '{field_name}' preenchido com sucesso.")
        return True
    except Exception as e:
        tratar_falha(driver_instance, f"preencher_campo_{field_name}", e)
        return False

def clicar_botao_pesquisar(driver_instance):
    try:
        logging.info("Tentando localizar e clicar no botão 'Pesquisar'...")
        botao_pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'js-searchButton') and text()='Pesquisar']")))
        botao_pesquisar.click()
        logging.info("Botão 'Pesquisar' clicado com sucesso!")
        return True
    except Exception as e:
        tratar_falha(driver_instance, "clicar_botao_pesquisar", e)
        return False

def selecionar_e_baixar_documentos(driver_instance, table_id="primaryHitlist_grid"):
    # ETAPA 1: SELECIONAR E COLETAR DADOS DA TABELA
    docs_para_enviar = []
    try:
        logging.info(f"Selecionando e coletando dados da tabela '{table_id}'...")
        tabela = wait.until(EC.visibility_of_element_located((By.ID, table_id)))
        linhas = tabela.find_elements(By.XPATH, ".//tbody/tr[@role='row']")
        if not linhas:
            logging.warning(f"Nenhuma linha de dados encontrada na tabela '{table_id}'.")
            return False

        for linha in linhas:
            driver_instance.execute_script(
                "arguments[0].setAttribute('aria-selected', 'true'); arguments[0].classList.add('ui-iggrid-activerow', 'ui-state-focus'); var cells = arguments[0].getElementsByTagName('td'); for (var j = 0; j < cells.length; j++) { cells[j].classList.add('ui-iggrid-selectedcell', 'ui-state-active'); }",
                linha
            )
            
            # CORREÇÃO: Lendo o ID da linha e assumindo os outros valores
            id_completo = linha.get_attribute('id')
            doc_id = id_completo.split('_')[-1]
            
            if doc_id:
                # Usando os valores de revid e filetypeid que vimos na sua captura de rede
                docs_para_enviar.append({"docid": doc_id, "revid": "-1", "filetypeid": "453"})
        
        logging.info(f"{len(docs_para_enviar)} documentos selecionados e dados coletados.")

    except Exception as e:
        tratar_falha(driver_instance, f"selecionar_e_coletar_dados", e)
        return False

    if not docs_para_enviar:
        logging.error("Nenhum documento com ID foi coletado para o download.")
        return False
        
    # ETAPA 2: ENVIAR REQUISIÇÃO DE REDE E BAIXAR
    try:
        logging.info("Iniciando processo de download via requisição POST...")
        initial_window_handle = driver_instance.current_window_handle
        cookies = driver_instance.get_cookies()
        session_cookies = {c['name']: c['value'] for c in cookies}
        
        url_post = "https://im40333.onbaseonline.com/203appnet/SendToFile.aspx"
        payload = {
            "Action": "SendTo",
            "SendToOption": "File",
            "selectedDocs": json.dumps(docs_para_enviar)
        }
        logging.info(f"Payload preparado para envio: {payload}")

        response = requests.post(url_post, cookies=session_cookies, data=payload)
        response.raise_for_status()
        
        popup_html = response.text
        if "Salvar em arquivo" not in popup_html:
            raise Exception("A resposta do servidor não contém a página de popup esperada.")
        
        logging.info("Resposta do popup recebida com sucesso. Abrindo em nova aba...")
        driver_instance.switch_to.new_window('tab')
        driver_instance.get("data:text/html;charset=utf-8," + popup_html)

        logging.info("Interagindo com os elementos na nova aba...")
        wait.until(EC.element_to_be_clickable((By.ID, "CheckboxExportNote"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "btnSave"))).click()
        logging.info("Opções configuradas e 'Salvar' clicado.")
        
        time.sleep(5)
        driver_instance.close()
        driver_instance.switch_to.window(initial_window_handle)
        
        start_time = time.time()
        while time.time() - start_time < 60:
            files = [f for f in os.listdir(download_directory) if not f.endswith(('.tmp', '.crdownload'))]
            if files:
                logging.info(f"Download do arquivo '{files[0]}' iniciado com sucesso!")
                return True
            time.sleep(1)
        logging.warning("Nenhum novo arquivo detectado na pasta de download.")
        return False

    except Exception as e:
        content_para_salvar = response.text if 'response' in locals() else ""
        tratar_falha(driver_instance, "envio_requisicao_post", e, content_para_salvar)
        return False
    finally:
        driver_instance.switch_to.default_content()

# ========== FLUXO PRINCIPAL DA AUTOMAÇÃO ==========
try:
    url = "https://im40333.onbaseonline.com/203appnet/Login.aspx"
    logging.info(f"Acessando: {url}")
    driver.get(url)
    
    try:
        logging.info("Localizando campos de login e senha...")
        campo_login = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/input")))
        campo_senha = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div/div[2]/table/tbody/tr[3]/td[2]/input")))
        botao_login = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[3]/div/div[2]/table/tbody/tr[5]/td[2]/button")))
        campo_login.send_keys("60036544")
        campo_senha.send_keys("12345678")
        botao_login.click()
        logging.info("Login realizado com sucesso.")
    except Exception as e:
        tratar_falha(driver, "etapa_login", e)
        raise 

    try:
        logging.info("Aguardando e mudando para o iframe 'NavPanelIFrame'...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "NavPanelIFrame")))
        campo_anexos = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div[1]/div/div/div/div[1]/div//input")))
        campo_anexos.send_keys("ANEXOS")
        time.sleep(1)
        botao_label = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[2]/div[1]/div/div/div/div[3]/div[1]/ul/li[5]/label")))
        botao_label.click()
        logging.info("Interação com 'ANEXOS' concluída com sucesso.")
    except Exception as e:
        tratar_falha(driver, "interacao_anexos_iframe", e)
        raise
    
    if preencher_campo_rh(driver, "114", "2469189", "RH - Matricula"):
        if clicar_botao_pesquisar(driver):
            driver.switch_to.default_content()
            time.sleep(15)
            try:
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frmViewer")))
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameDocSelect")))
                if selecionar_e_baixar_documentos(driver):
                    logging.info("FLUXO COMPLETO EXECUTADO COM SUCESSO!")
                else:
                    logging.error("Status Final: Falha na etapa de download.")
            except Exception as e:
                 tratar_falha(driver, "navegacao_iframes_de_resultado", e)
        else:
            logging.error("Status Final: Botão 'Pesquisar' não foi clicado.")
    else:
        logging.error("Status Final: Campo de matrícula não foi preenchido.")

    logging.info("Automação concluída. Pressione ESC para encerrar o navegador.")
    while True:
        if keyboard.is_pressed("esc"):
            logging.info("ESC pressionado. Encerrando o navegador...")
            break
        time.sleep(0.1)

except Exception as e:
    tratar_falha(driver, "ERRO_FATAL_FLUXO_PRINCIPAL", e)

finally:
    if 'driver' in locals() and driver:
        driver.quit()
    logging.info("Navegador encerrado com sucesso.")