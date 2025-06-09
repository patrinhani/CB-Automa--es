from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import keyboard
import time
import logging
from datetime import datetime
import os

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

# ============ CONFIGURAÇÕES DE DOWNLOAD (para Selenium) ============
download_directory = r"C:\Users\2160036544\Downloads\Automacao_Arquivos"
if not os.path.exists(download_directory):
    os.makedirs(download_directory)
    logging.info(f"Diretório de download criado: {download_directory}")
prefs = {
    "download.default_directory": download_directory,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
options.add_experimental_option("excludeSwitches", ["disable-popup-blocking"])

logging.info("Inicializando o Edge WebDriver...")
service = Service(driver_path)
driver = webdriver.Edge(service=service, options=options)
wait = WebDriverWait(driver, 40)

# ========== LÓGICA DE TRATAMENTO DE FALHAS ==========
def salvar_html_para_depuracao(driver_instance, filename_prefix="debug_html"):
    try:
        driver_instance.switch_to.default_content()
        html_content = driver_instance.page_source
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_{current_time}.html"
        filepath = os.path.join(os.getcwd(), filename)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(html_content)
        logging.info(f"HTML da falha salvo em: {filepath}")
    except Exception as e:
        logging.error(f"FALHA CRÍTICA ao salvar o HTML de depuração: {e}", exc_info=True)

def salvar_screenshot_para_depuracao(driver_instance, filename_prefix="debug_screenshot"):
    try:
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_{current_time}.png"
        filepath = os.path.join(os.getcwd(), filename)
        driver_instance.save_screenshot(filepath)
        return filepath
    except Exception as e:
        logging.error(f"FALHA ao salvar screenshot: {e}", exc_info=True)
        return None

def tratar_falha_e_salvar_html(driver_instance, contexto_da_falha, exception_obj):
    mensagem_erro = f"FALHA no contexto '{contexto_da_falha}': {exception_obj}"
    logging.error(mensagem_erro, exc_info=True)
    nome_arquivo_falha = f"FALHA_{contexto_da_falha.replace(' ', '_').replace('-', '_')}"
    salvar_html_para_depuracao(driver_instance, nome_arquivo_falha)
    salvar_screenshot_para_depuracao(driver_instance, nome_arquivo_falha)

# ========== FUNÇÕES DE AUTOMAÇÃO (DO SEU CÓDIGO BASE) ==========
def preencher_campo_rh(driver_instance, keyword_type_id, value_to_send, field_name):
    try:
        logging.info(f"Tentando preencher o campo '{field_name}'...")
        parent_div = wait.until(EC.visibility_of_element_located((By.XPATH, f"//div[@keywordtypeid='{keyword_type_id}']")))
        campo_input = wait.until(lambda d: parent_div.find_element(By.XPATH, ".//input[contains(@class, 'keywordInput')]"))
        wait.until(EC.element_to_be_clickable(campo_input))
        campo_input.clear()
        campo_input.send_keys(value_to_send)
        logging.info(f"Campo '{field_name}' preenchido com sucesso.")
        return True
    except Exception as e:
        tratar_falha_e_salvar_html(driver_instance, f"preencher_campo_{field_name}", e)
        return False

def clicar_botao_pesquisar(driver_instance):
    try:
        logging.info("Tentando localizar e clicar no botão 'Pesquisar'...")
        botao_pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'js-searchButton') and text()='Pesquisar']")))
        botao_pesquisar.click()
        logging.info("Botão 'Pesquisar' clicado com sucesso!")
        return True
    except Exception as e:
        tratar_falha_e_salvar_html(driver_instance, "clicar_botao_pesquisar", e)
        return False

def selecionar_todos_itens_tabela(driver_instance, table_id="primaryHitlist_grid"):
    try:
        logging.info(f"Tentando selecionar todos os itens da tabela '{table_id}'...")
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
        logging.info(f"Todas as {len(linhas)} linhas foram selecionadas via JavaScript.")
        return True
    except Exception as e:
        tratar_falha_e_salvar_html(driver_instance, f"selecionar_itens_tabela_{table_id}", e)
        return False

### INÍCIO DA SEÇÃO MODIFICADA ###
def enviar_para_arquivo_e_baixar(driver_instance):
    try:
        logging.info("Iniciando processo de download com construção de URL...")
        initial_window_handle = driver_instance.current_window_handle
        
        # 1. Obter o OBToken da página atual, que é necessário para a URL do popup
        token = driver_instance.execute_script("return window.Page.Get('__OBToken');")
        if not token:
            raise Exception("Não foi possível obter o OBToken da página.")
        logging.info(f"OBToken obtido: {token}")

        # 2. Construir a URL do popup de download diretamente
        popup_url = f"https://im40333.onbaseonline.com/203appnet/SendToFile.aspx?OBToken={token}"
        logging.info(f"URL do popup construída: {popup_url}")

        # 3. Abrir a URL construída em uma nova aba
        logging.info("Abrindo URL do popup em uma nova aba...")
        driver_instance.switch_to.new_window('tab')
        driver_instance.get(popup_url)

        # 4. Interagir com os elementos na nova aba (o "popup")
        logging.info("Interagindo com os elementos na nova aba...")
        wait.until(EC.element_to_be_clickable((By.ID, "CheckboxExportNote"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "btnSave"))).click()
        logging.info("Opções configuradas e 'Salvar' clicado na nova aba.")
        
        # 5. Esperar o início do download, fechar a aba e voltar para a principal
        time.sleep(3) 
        driver_instance.close()
        driver_instance.switch_to.window(initial_window_handle)
        
        # 6. Verificação do arquivo baixado
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
        tratar_falha_e_salvar_html(driver_instance, "construcao_url_e_download", e)
        return False
    finally:
        # Garante que o foco sempre volte para a janela/aba principal
        handles = driver_instance.window_handles
        if initial_window_handle in handles and driver_instance.current_window_handle != initial_window_handle:
            driver_instance.switch_to.window(initial_window_handle)
        driver_instance.switch_to.default_content()
### FIM DA SEÇÃO MODIFICADA ###


# ========== FLUXO PRINCIPAL DA AUTOMAÇÃO (SEU CÓDIGO ORIGINAL) ==========
try:
    url = "https://im40333.onbaseonline.com/203appnet/Login.aspx"
    logging.info(f"Acessando: {url}")
    driver.get(url)
    
    # Etapa 1: Login
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
        tratar_falha_e_salvar_html(driver, "etapa_login", e)
        raise 

    # Etapa 2: Interação com 'ANEXOS'
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
        tratar_falha_e_salvar_html(driver, "interacao_anexos_iframe", e)
        raise
    
    # Etapa 3: Preenchimento, Pesquisa e Download
    if preencher_campo_rh(driver, "114", "2469189", "RH - Matricula"):
        if clicar_botao_pesquisar(driver):
            driver.switch_to.default_content()
            time.sleep(15)
            try:
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frmViewer")))
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameDocSelect")))
                if selecionar_todos_itens_tabela(driver):
                    if enviar_para_arquivo_e_baixar(driver):
                        logging.info("FLUXO COMPLETO EXECUTADO COM SUCESSO!")
                    else:
                        logging.error("Status Final: Falha na etapa de download.")
                else:
                    logging.error("Status Final: Falha ao selecionar itens da tabela.")
            except Exception as e:
                 tratar_falha_e_salvar_html(driver, "navegacao_iframes_de_resultado", e)
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
    tratar_falha_e_salvar_html(driver, "ERRO_FATAL_FLUXO_PRINCIPAL", e)

finally:
    if 'driver' in locals() and driver:
        driver.quit()
    logging.info("Navegador encerrado com sucesso.")