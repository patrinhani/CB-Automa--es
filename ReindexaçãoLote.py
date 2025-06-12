import logging
import os
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Função de ajuda para depurar
def save_page_html(driver, step_name):
    os.makedirs("debug_logs", exist_ok=True)
    filename = os.path.join("debug_logs", f"html_error_at_{step_name}_{time.strftime('%Y%m%d-%H%M%S')}.html")
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        logging.info(f"HTML de depuração salvo para '{step_name}' em: {filename}")
    except Exception as e:
        logging.error(f"Erro ao salvar HTML de depuração: {e}")

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

def preencher_e_enviar_lote():
    caminho_perfil_edge = r"C:\Users\2160036544\Downloads\Demandas\teste\gedocReindexações"
    caminho_driver_edge = r"C:\Users\2160036544\Downloads\edgedriver_win64\msedgedriver.exe"
    url_site = "https://im40333.onbaseonline.com/203WPT/index.html?application=RH&locale=pt-BR"
    
    matriculas_para_adicionar = [
        "3367703", "2301679", "60008034", "60028569", "4075846", "60033972",
        "3931994", "1835254", "2697459", "724645", "4837380", "1268856",
        "4065484", "2038498", "1221078"
    ]

    if not os.path.exists(caminho_driver_edge):
        logging.error(f"ERRO: O msedgedriver.exe não foi encontrado em: {caminho_driver_edge}")
        return

    edge_options = EdgeOptions()
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument(f"user-data-dir={caminho_perfil_edge}")

    service = EdgeService(caminho_driver_edge)
    driver = None

    try:
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.get(url_site)
        logging.info(f"1. Acessando o site: {url_site}")

        ## --- ETAPA OTIMIZADA ---
        logging.info("2. Aguardando o painel de formulários e clicando no botão '+ Lote'...")
        seletor_otimizado_lote = "#formListPanel td[data-name='RH - Solicitacao de Lote']"
        botao_lote = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, seletor_otimizado_lote))
        )
        botao_lote.click()
        logging.info("   -> Botão '+ Lote' clicado.")
        ## -------------------------
        
        logging.info("3. Mudando para o primeiro iframe (ID: iframeHolder)...")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "iframeHolder"))
        )
        logging.info("   -> Foco alterado para 'iframeHolder'.")
        
        logging.info("4. Mudando para o iframe aninhado (ID: uf_hostframe)...")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "uf_hostframe"))
        )
        logging.info("   -> Foco alterado para 'uf_hostframe'.")

        logging.info(f"5. Iniciando o cadastro de {len(matriculas_para_adicionar)} matrículas...")
        
        botao_adicionar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[hsi-automation='automation-addItem']"))
        )
        
        for i, matricula_atual in enumerate(matriculas_para_adicionar):
            logging.info(f"   -> Adicionando linha {i+1} para a matrícula '{matricula_atual}'")
            botao_adicionar.click()
            id_campo_dinamico = f"rhmatricula12_input_{i}"
            campo_matricula = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, id_campo_dinamico))
            )
            campo_matricula.send_keys(matricula_atual)
            logging.info(f"      -> Campo preenchido com sucesso.")
            time.sleep(0.5)
        
        logging.info(f"   -> Todas as {len(matriculas_para_adicionar)} matrículas foram adicionadas.")

        logging.info("6. Clicando no botão 'Enviar' para finalizar o processo...")
        botao_enviar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='Enviar']"))
        )
        botao_enviar.click()
        logging.info("   -> Botão 'Enviar' clicado.")
        
        logging.info("--- PROCESSO CONCLUÍDO: Matrículas enviadas com sucesso! ---")
        logging.info("O navegador será fechado em 5 segundos...")
        time.sleep(5)

    except TimeoutException:
        logging.error("ERRO: Um elemento (botão ou campo) não foi encontrado a tempo.")
        if driver: save_page_html(driver, "timeout_error")
        input("Ocorreu um erro. Pressione Enter para fechar...") 
    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado: {e}")
        if driver: save_page_html(driver, "unexpected_error")
        input("Ocorreu um erro. Pressione Enter para fechar...")
    finally:
        if driver:
            driver.quit()
            logging.info("Navegador fechado.")

if __name__ == "__main__":
    preencher_e_enviar_lote()