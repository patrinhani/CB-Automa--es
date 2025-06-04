from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import keyboard
import time
import logging

# ========== CONFIGURAÇÃO DO LOG ==========
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S"
)

# ========== CONFIGURAÇÃO DO EDGE DRIVER ==========
driver_path = r"C:\Users\2160036544\Downloads\edgedriver_win64\msedgedriver.exe"
options = Options()
# options.add_argument("--headless") # Descomente para rodar o navegador em modo invisível
# options.add_argument("--disable-gpu") # Necessário para headless em alguns sistemas
# options.add_argument("--no-sandbox") # Essencial para ambientes Linux/Docker

logging.info("Inicializando o Edge WebDriver...")
service = Service(driver_path)
driver = webdriver.Edge(service=service, options=options)
wait = WebDriverWait(driver, 15) # Aumentado o tempo de espera para 15 segundos

# ========== FUNÇÃO PARA PREENCHER CAMPOS DE RH (KEYWORDTYPEID) ==========
def preencher_campo_rh(driver_instance, keyword_type_id, value_to_send, field_name):
    """
    Tenta localizar um campo de entrada de RH pelo seu keywordtypeid
    e preenche com o valor fornecido.

    Args:
        driver_instance: A instância do WebDriver.
        keyword_type_id (str): O valor do atributo 'keywordtypeid' do div pai.
        value_to_send (str): O texto a ser enviado para o campo de entrada.
        field_name (str): O nome amigável do campo para mensagens de log.
    Returns:
        bool: True se o campo foi preenchido com sucesso, False caso contrário.
    """
    try:
        logging.info(f"Tentando preencher o campo '{field_name}' (keywordtypeid='{keyword_type_id}')...")
        
        # 1. Tenta localizar o div pai
        logging.debug(f"Buscando o div pai com keywordtypeid='{keyword_type_id}'...")
        parent_div = wait.until(EC.presence_of_element_located((By.XPATH, f"//div[@keywordtypeid='{keyword_type_id}']")))
        logging.debug(f"Div pai para '{field_name}' encontrado.")
        
        # 2. Tenta localizar o input dentro do div pai
        logging.debug(f"Buscando o input dentro do div pai para '{field_name}'...")
        campo_input = wait.until(EC.element_to_be_clickable((parent_div, By.XPATH, ".//input[contains(@class, 'keywordInput')]")))
        logging.debug(f"Input para '{field_name}' encontrado e clicável.")
        
        campo_input.clear() # Limpa qualquer texto existente
        campo_input.send_keys(value_to_send)
        logging.info(f"Campo '{field_name}' preenchido com sucesso: '{value_to_send}'.")
        return True
    except Exception as e:
        logging.error(f"FALHA ao preencher o campo '{field_name}' (keywordtypeid='{keyword_type_id}'): {e}")
        logging.debug(f"Detalhes da exceção ao preencher '{field_name}':", exc_info=True) # Adiciona stack trace
        return False

try:
    # ========== ACESSA A PÁGINA DE LOGIN ==========
    url = "https://im40333.onbaseonline.com/203appnet/Login.aspx"
    logging.info(f"Acessando: {url}")
    driver.get(url)
    time.sleep(2) 

    # ========== OBTÉM E MOSTRA O TÍTULO ==========
    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.string if soup.title else "Título não encontrado"
    logging.info(f"Título da página: {title}")

    # ========== REALIZA O LOGIN ==========
    logging.info("Localizando campos de login e senha...")
    campo_login = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/input")))
    campo_senha = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div/div[2]/table/tbody/tr[3]/td[2]/input")))
    botao_login = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[3]/div/div[2]/table/tbody/tr[5]/td[2]/button")))

    logging.info("Preenchendo credenciais...")
    campo_login.send_keys("60036544")
    campo_senha.send_keys("12345678")

    logging.info("Clicando no botão de login...")
    botao_login.click()

    # ========== ESPERA PÓS-LOGIN E TROCA DE IFRAME PARA 'ANEXOS' ==========
    time.sleep(5) # Aumentei a espera para 5 segundos após o login

    logging.info("Verificando iframes para digitar 'ANEXOS' e clicar na label.")
    anexos_interagido = False

    # Sempre volta para o default_content antes de buscar iframes novamente
    driver.switch_to.default_content() 
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    logging.info(f"Quantidade de iframes na página: {len(iframes)}")

    for i, iframe in enumerate(iframes):
        logging.info(f"Tentando trocar para iframe {i} (para 'ANEXOS')...")
        try:
            driver.switch_to.frame(iframe)
            logging.debug(f"Sucesso ao trocar para iframe {i} (para 'ANEXOS').")
            
            # Tenta localizar e preencher o input 'ANEXOS'
            campo_anexos = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div[1]/div/div/div/div[1]/div//input")))
            campo_anexos.send_keys("ANEXOS")
            logging.info("Texto 'ANEXOS' inserido com sucesso no iframe.")
            time.sleep(1) 

            # Tenta clicar na label do 5º item da lista (referente a 'ANEXOS')
            botao_label = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[2]/div[1]/div/div/div/div[3]/div[1]/ul/li[5]/label")))
            botao_label.click()
            logging.info("Elemento (label de 'ANEXOS') clicado com sucesso no iframe.")
            anexos_interagido = True
            break 
        except Exception as e:
            logging.debug(f"Não encontrou os elementos de 'ANEXOS' dentro do iframe {i}: {e}")
        finally:
            driver.switch_to.default_content() # Sempre volta para o conteúdo principal

    if not anexos_interagido:
        logging.warning("Não foi possível localizar e interagir com os elementos 'ANEXOS' em nenhum iframe. Tentando na página principal.")
        try:
            campo_anexos = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/div[1]/div/div/div/div[1]/div//input")))
            campo_anexos.send_keys("ANEXOS")
            logging.info("Texto 'ANEXOS' inserido diretamente (fora de iframe).")
            time.sleep(1)

            botao_label = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[2]/div[1]/div/div/div/div[3]/div[1]/ul/li[5]/label")))
            botao_label.click()
            logging.info("Elemento (label de 'ANEXOS') clicado diretamente (fora de iframe).")
            anexos_interagido = True
        except Exception as e:
            logging.warning(f"Não foi possível localizar e interagir com os elementos 'ANEXOS' na página principal: {e}")

    if not anexos_interagido:
        logging.error("Não foi possível concluir a interação com 'ANEXOS'. A automação de preenchimento dos campos de RH pode falhar.")
    
    # ========== ESPERA CRÍTICA PÓS-INTERAÇÃO COM 'ANEXOS' ==========
    # Aumentado o tempo para garantir que os campos de RH carreguem.
    logging.info("Aguardando carregamento dos campos de RH após seleção de 'ANEXOS'...")
    time.sleep(7) # Aumentei significativamente esta espera para 7 segundos

    # ========== PREENCHER CAMPOS DE RH (APÓS INTERAÇÃO COM 'ANEXOS') ==========
    logging.info("Iniciando o preenchimento dos campos de RH (Matrícula, CPF, etc.)...")
    campos_rh_preenchidos = False

    # É FUNDAMENTAL garantir que estamos no conteúdo principal antes de buscar NOVOS iframes.
    driver.switch_to.default_content() 
    iframes_pos_anexos = driver.find_elements(By.TAG_NAME, "iframe")
    
    if iframes_pos_anexos:
        logging.info(f"Encontrados {len(iframes_pos_anexos)} iframe(s) após a interação com 'ANEXOS'. Tentando preencher campos de RH dentro deles.")
        for i, iframe in enumerate(iframes_pos_anexos):
            logging.info(f"Tentando trocar para iframe {i} para preencher campos de RH...")
            try:
                driver.switch_to.frame(iframe)
                logging.debug(f"Sucesso ao trocar para iframe {i} para campos de RH.")

                # Tenta preencher os campos de RH
                # ATUALIZADO: Usando a matrícula fornecida
                matricula_ok = preencher_campo_rh(driver, "114", "2469189", "RH - Matricula")
                
                # Se a matrícula foi preenchida, tentamos os outros.
                # Se a matrícula falhou, os outros provavelmente falharão também.
                if matricula_ok:
                    campos_rh_preenchidos = True
                    preencher_campo_rh(driver, "118", "98765432100", "RH - CPF")
                    preencher_campo_rh(driver, "119", "EMPRESA TESTE LTDA", "RH - Empresa")
                    preencher_campo_rh(driver, "113", "PROCESSO AUTOMATICO 123", "RH - Processo")
                    preencher_campo_rh(driver, "133", "PROT-456789", "RH - Numero Protocolo")
                else:
                    logging.warning(f"Matrícula não preenchida no iframe {i}, pulando outros campos de RH neste iframe.")

                if campos_rh_preenchidos:
                    logging.info("Pelo menos um campo de RH foi preenchido dentro do iframe pós-'ANEXOS'.")
                    break # Sai do loop de iframes se os campos forem preenchidos
            except Exception as e:
                # Captura de exceção mais genérica para evitar que um erro em um iframe bloqueie os outros
                logging.debug(f"Erro inesperado ao tentar interagir com iframe {i} para campos de RH: {e}")
            finally:
                driver.switch_to.default_content() # Sempre volta para o conteúdo principal

    if not campos_rh_preenchidos:
        logging.info("Não foi possível preencher campos de RH dentro de iframes pós-'ANEXOS'. Tentando na página principal (fora de iframes).")
        # Tenta preencher na página principal se não encontrou em iframes
        matricula_ok = preencher_campo_rh(driver, "114", "2469189", "RH - Matricula")
        if matricula_ok:
            campos_rh_preenchidos = True
            preencher_campo_rh(driver, "118", "98765432100", "RH - CPF")
            preencher_campo_rh(driver, "119", "EMPRESA TESTE LTDA", "RH - Empresa")
            preencher_campo_rh(driver, "113", "PROCESSO AUTOMATICO 123", "RH - Processo")
            preencher_campo_rh(driver, "133", "PROT-456789", "RH - Numero Protocolo")
        else:
            logging.warning("Matrícula não preenchida na página principal, pulando outros campos de RH.")

    if not campos_rh_preenchidos:
        logging.error("NÃO FOI POSSÍVEL LOCALIZAR E PREENCHER A MATRÍCULA OU QUALQUER OUTRO CAMPO DE RH.")
    
    # ========== LOOP DE ESPERA PARA INTERAÇÃO MANUAL ==========
    logging.info("Automação inicial concluída. Pressione ESC para encerrar o navegador.")
    while True:
        if keyboard.is_pressed("esc"):
            logging.info("ESC pressionado. Encerrando o navegador...")
            break
        time.sleep(0.1)

except Exception as e:
    logging.error(f"Ocorreu um erro inesperado durante a automação principal: {e}", exc_info=True)

finally:
    if driver:
        driver.quit()
        logging.info("Navegador encerrado com sucesso.")