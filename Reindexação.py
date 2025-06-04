from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import Select 
import logging
import time
import os

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

# Definir um diretório para salvar logs de depuração
DEBUG_DIR = "selenium_debug_logs"
os.makedirs(DEBUG_DIR, exist_ok=True)


def injetar_override_window_open(driver):
    """
    Injeta um script JavaScript para interceptar a URL de popups abertas por window.open().
    A URL capturada é armazenada em window.openedURL.
    """
    logging.info("Injetando override em window.open() para capturar URL de popup...")
    driver.execute_script("window.openedURL = null;") 
    driver.execute_script("""
        window._originalWindowOpen = window.open;
        window.open = function(url, name, specs) {
            if (url && url !== 'about:blank' && url !== '') {
                window.openedURL = url;
                console.log('Intercepted window.open URL:', url); 
            }
            return window._originalWindowOpen(url, name, specs);
        };
    """)


def capturar_url_popup(driver):
    """
    Recupera a URL da popup que foi capturada pelo override de window.open().
    Espera até que a URL esteja presente ou atinge um timeout.
    """
    logging.info("Aguardando captura da URL da popup...")
    try:
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return window.openedURL;") is not None and \
                      d.execute_script("return window.openedURL;") != '' and \
                      d.execute_script("return window.openedURL;") != 'about:blank'
        )
        url = driver.execute_script("return window.openedURL;")
        logging.info(f"URL capturada pelo override de window.open: {url}")
        driver.execute_script("window.openedURL = null;") # Resetar após a captura
        return url
    except TimeoutException:
        logging.warning("Timeout: Nenhuma URL de popup foi capturada dentro do tempo esperado.")
        return None


def clicar_no_botao(driver):
    """
    Aguarda o botão 'Buscar' ficar clicável e clica nele.
    """
    logging.info("Aguardando botão 'Buscar' ficar clicável e clicando...")
    btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "keywordComponentApplyEditBtn"))
    )
    btn.click()


def abrir_url_manual(driver, url):
    """
    Abre uma nova aba no navegador e navega para a URL fornecida.
    Espera até que o corpo da página esteja presente para garantir o carregamento inicial.
    """
    logging.info(f"Abrindo URL manualmente em nova aba: {url}")
    driver.execute_script("window.open('');") 
    driver.switch_to.window(driver.window_handles[-1]) # Muda o foco para a nova aba (a última aberta)
    driver.get(url) # Navega para a URL na nova aba
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
    logging.info("Nova aba carregada com sucesso.")


def selecionar_e_desselecionar_linhas(driver):
    """
    Percorre todas as linhas da tabela. Tenta selecionar as linhas "RH - ANEXOS"
    simulando um clique JavaScript. Desseleciona qualquer outra linha.
    Esta função assume que o foco do Selenium JÁ ESTÁ no iframe correto.
    """
    logging.info("Iniciando seleção e desseleção de linhas na tabela...")

    try:
        # Espera que a tabela esteja visível e que as linhas com data-id existam
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "//tr[@role='row'][@data-id]"))
        )
        logging.info("Pelo menos uma linha da tabela está visível. Aguardando um momento para carregar todas.")
        time.sleep(4) # Uma pequena pausa para garantir que todos os dados sejam renderizados

        todas_as_linhas = driver.find_elements(By.XPATH, "//tr[@role='row'][@data-id]")
        logging.info(f"Encontradas {len(todas_as_linhas)} linhas na tabela para processar.")

        if not todas_as_linhas:
            logging.warning("Nenhuma linha encontrada na tabela. Verifique se os dados carregaram ou o XPath para as linhas.")
            return

        for i, row in enumerate(todas_as_linhas):
            try:
                doc_type_cell = row.find_element(
                    By.XPATH,
                    "./td[@aria-describedby='primaryHitlist_grid_DocumentTypeName']"
                )
                doc_type_text = doc_type_cell.text.strip()
                row_data_id = row.get_attribute('data-id')
                
                # Verifica o estado atual de seleção (apenas para logging)
                is_selected = row.get_attribute('aria-selected') == 'true'
                
                if doc_type_text == "RH - ANEXOS":
                    # Tenta forçar a seleção da linha via clique JS
                    # Se o site gerencia seleção múltipla sem CTRL, um clique simples em cada uma basta.
                    if not is_selected: # Clica apenas se não estiver já selecionado
                        driver.execute_script("arguments[0].click();", row)
                        logging.info(f"Linha {i+1} (data-id: {row_data_id}, Tipo: '{doc_type_text}') clicada para seleção.")
                    else:
                        logging.info(f"Linha {i+1} (data-id: {row_data_id}, Tipo: '{doc_type_text}') já estava selecionada.")
                    
                    # Verificação adicional após o clique para garantir a seleção (pode ser útil para depuração)
                    time.sleep(0.1) # Pequena pausa para o clique processar
                    if row.get_attribute('aria-selected') != 'true':
                        logging.warning(f"Atenção: Linha {i+1} ('{doc_type_text}') não permaneceu selecionada após o clique JS.")
                        # Tenta uma re-seleção mais forçada se a anterior falhou
                        driver.execute_script("arguments[0].setAttribute('aria-selected', 'true'); arguments[0].classList.add('ui-iggrid-selectedrow');", row)
                        cells_in_row = row.find_elements(By.TAG_NAME, "td")
                        for cell in cells_in_row:
                            driver.execute_script("arguments[0].classList.add('ui-iggrid-selectedcell', 'ui-state-active');", cell)
                        logging.info(f"Tentativa de forçar seleção de Linha {i+1} via setAttribute/classList.")

                else:
                    # Desseleciona explicitamente outras linhas se estiverem selecionadas
                    if is_selected:
                        driver.execute_script("arguments[0].click();", row) # Tenta um clique para desselecionar
                        logging.info(f"Linha {i+1} (data-id: {row_data_id}, Tipo: '{doc_type_text}') clicada para desseleção.")
                    else:
                        # Se não estava selecionada, apenas garante que os atributos estejam corretos
                        driver.execute_script("arguments[0].setAttribute('aria-selected', 'false'); arguments[0].classList.remove('ui-iggrid-selectedrow');", row)
                        cells_in_row = row.find_elements(By.TAG_NAME, "td")
                        for cell in cells_in_row:
                            driver.execute_script("arguments[0].classList.remove('ui-iggrid-selectedcell', 'ui-state-active');", cell)
                        logging.info(f"Linha {i+1} (data-id: {row_data_id}, Tipo: '{doc_type_text}') desselecionada (ou já estava).")
                
                time.sleep(0.2) # Pausa entre o processamento das linhas para evitar sobrecarga

            except Exception as row_process_error:
                logging.error(f"Falha ao processar a linha {i+1} (data-id: {row.get_attribute('data-id')}): {row_process_error}")
                timestamp = time.strftime("%Y%m%d-%H%M%S")
                error_html_path = os.path.join(DEBUG_DIR, f"error_html_row_error_{timestamp}_item_{i+1}.html")
                try:
                    with open(error_html_path, "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    logging.info(f"HTML da página salvo em: {error_html_path}")
                except Exception as save_error:
                    logging.error(f"Erro ao tentar salvar HTML de depuração: {save_error}")
                raise

        logging.info("Processo de seleção/desseleção de linhas concluído.")
        
        # Verificação final de quantas linhas RH - ANEXOS estão realmente selecionadas
        final_selected_rh_anexos_count = driver.execute_script("""
            let count = 0;
            document.querySelectorAll("tr[role='row'][data-id]").forEach(row => {
                let docTypeCell = row.querySelector("td[aria-describedby='primaryHitlist_grid_DocumentTypeName']");
                if (docTypeCell && docTypeCell.textContent.trim() === 'RH - ANEXOS' && row.getAttribute('aria-selected') === 'true') {
                    count++;
                }
            });
            return count;
        """)
        logging.info(f"Total de linhas 'RH - ANEXOS' finalmente selecionadas: {final_selected_rh_anexos_count}")


    except Exception as e:
        logging.error(f"Erro geral durante a seleção e desseleção de linhas: {e}")
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        error_html_path = os.path.join(DEBUG_DIR, f"error_html_global_selection_error_{timestamp}.html")
        try:
            with open(error_html_path, "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logging.info(f"HTML da página salvo em: {error_html_path}")
        except Exception as save_error:
            logging.error(f"Erro ao tentar salvar HTML de depuração global: {save_error}")
        raise


def interagir_com_modal_salvar_arquivo(driver):
    """
    Interage com o modal/janela que aparece após clicar em 'Arquivo' no menu de contexto.
    Isso inclui selecionar 'PDF' e clicar em 'Salvar'.
    """
    logging.info("Iniciando interação com o modal 'Salvar em Arquivo'...")
    try:
        driver.switch_to.default_content() 
        logging.info("Foco do Selenium mudado para o conteúdo padrão da janela para interagir com o modal.")

        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, "frmSaveToFile"))
        )
        logging.info("Modal 'Salvar em Arquivo' detectado e visível.")
        time.sleep(1) 

        logging.info("Selecionando 'PDF (.pdf)' no dropdown 'Salvar como'...")
        select_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "selectContent"))
        )
        select_obj = Select(select_element)
        select_obj.select_by_value("pdf") 
        logging.info("Opção 'PDF (.pdf)' selecionada.")
        time.sleep(1) 

        logging.info("Clicando no botão 'Salvar'...")
        save_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnSave"))
        )
        save_button.click()
        logging.info("Botão 'Salvar' clicado.")
        time.sleep(5) 

        try:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, "frmSaveToFile"))
            )
            logging.info("Modal 'Salvar em Arquivo' desapareceu.")
        except TimeoutException:
            logging.warning("Modal 'Salvar em Arquivo' não desapareceu em 10 segundos, pode estar processando ou necessitar de outra ação.")

        # Retorna o foco para o iframe principal após a interação com o modal
        driver.switch_to.default_content() 
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "DocSelectPage")))
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameDocSelect")))
        logging.info("Foco do Selenium retornado para o iframe 'frameDocSelect' após interagir com o modal.")


    except Exception as e:
        logging.error(f"Erro ao interagir com o modal 'Salvar em Arquivo': {e}")
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        error_html_path = os.path.join(DEBUG_DIR, f"error_html_save_file_modal_error_{timestamp}.html")
        try:
            with open(error_html_path, "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logging.info(f"HTML do modal salvo em: {error_html_path}")
        except Exception as save_error:
            logging.error(f"Erro ao tentar salvar HTML de depuração: {save_error}")
        raise


### **Função de Interação com Menu de Contexto (Otimizada e Única)**

def interagir_menu_contexto_salvar_arquivo(driver):
    """
    Executa o clique com o botão direito, simula hover e chama o handler JS
    para a opção 'Arquivo' do menu de contexto.
    """
    logging.info("Iniciando interação com o menu de contexto (Solução Otimizada)...")
    original_all_handles = driver.window_handles 

    try:
        # 1. Encontra a primeira célula "RH - ANEXOS" selecionada para o clique direito
        # (É importante que pelo menos uma esteja selecionada para o context_click,
        # mas as demais serão selecionadas via loop em selecionar_e_desselecionar_linhas)
        first_selected_rh_anexos_cell = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH, 
                "//tr[@role='row'][@data-id][@aria-selected='true']"
                "/td[@aria-describedby='primaryHitlist_grid_DocumentTypeName'"
                " and normalize-space(text())='RH - ANEXOS']"
            ))
        )
        logging.info(f"Primeira célula 'RH - ANEXOS' selecionada encontrada: {first_selected_rh_anexos_cell.text}")

        # 2. Dispara o evento 'contextmenu' via JavaScript na *célula da linha selecionada*.
        # Isso abre o menu de contexto.
        driver.execute_script("arguments[0].dispatchEvent(new Event('contextmenu', { bubbles: true }));", first_selected_rh_anexos_cell)
        logging.info("Evento 'contextmenu' disparado via JavaScript na linha selecionada.")
        
        # 3. Espera pelo item "Enviar para" (ID: menuControl_25) no menu de contexto principal.
        enviar_para_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "menuControl_25"))
        )
        logging.info("Item 'Enviar para' do menu de contexto (ID: menuControl_25) visível e clicável.")
        
        # 4. Simula o hover em "Enviar para" disparando um evento mouseover via JavaScript.
        # Isso é crucial para que o submenu "Arquivo" apareça, pois "Enviar para" não tem handler direto.
        driver.execute_script("arguments[0].dispatchEvent(new Event('mouseover', { bubbles: true }));", enviar_para_item)
        logging.info("Evento mouseover disparado em 'Enviar para' via JavaScript (para hover).")
        
        # Pequena pausa para garantir que o submenu apareça e estabilize.
        time.sleep(1.5) # Aumentei um pouco a pausa aqui para garantir o submenu.

        # 5. Espera que o objeto JavaScript 'window.contextMenu' esteja disponível
        # e tenha o método 'onMenuItemSelected'. Isso é uma forma robusta de garantir
        # que o controle de menu JavaScript esteja totalmente inicializado.
        WebDriverWait(driver, 10).until(
            lambda dr: dr.execute_script("return window.contextMenu && typeof window.contextMenu.onMenuItemSelected === 'function';")
        )
        logging.info("Objeto JavaScript 'window.contextMenu' e método 'onMenuItemSelected' detectados.")

        # 6. Chama diretamente o handler JavaScript para o item 'Arquivo' usando seu ID ('13').
        # Conforme o HTML, "Arquivo" tem um handler chamado "OnFile".
        # O `window.contextMenu.onMenuItemSelected('13')` é a forma correta de ativá-lo internamente.
        logging.info("Chamando diretamente o handler JavaScript 'onMenuItemSelected' para o item 'Arquivo' (ID 13).")
        driver.execute_script("window.contextMenu.onMenuItemSelected('13');")
        logging.info("Handler para 'Arquivo' executado via JavaScript.")
        
        # Limpa qualquer URL capturada anteriormente antes de verificar por nova janela.
        driver.execute_script("window.openedURL = null;")
        logging.info("window.openedURL resetado antes de verificar nova janela.")

        # 7. Verifica se uma nova janela/guia foi aberta (onde o modal de salvar pode estar)
        logging.info("Verificando se uma nova janela/guia foi aberta...")
        try:
            # Espera que o número de janelas aumente, indicando uma nova janela/tab do modal
            WebDriverWait(driver, 15).until(EC.number_of_windows_to_be(len(original_all_handles) + 1))
            new_window_handle = [handle for handle in driver.window_handles if handle not in original_all_handles][0]
            driver.switch_to.window(new_window_handle)
            logging.info(f"Nova janela/guia detectada com URL: {driver.current_url}")
            return True 
        except TimeoutException:
            # Se não abriu nova janela, o modal pode ter aparecido na mesma aba.
            logging.warning("Timeout: Nenhuma nova janela/guia foi detectada em 15 segundos após a execução do handler 'Arquivo'. Assumindo que o modal está na mesma aba.")
            logging.info(f"URL atual após ação do menu: {driver.current_url}")
            return True # Retorna True para que a função interagir_com_modal_salvar_arquivo seja chamada.
        except Exception as win_handle_error:
            logging.error(f"Erro ao verificar nova janela/guia: {win_handle_error}")
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            error_html_path = os.path.join(DEBUG_DIR, f"error_html_new_window_check_error_{timestamp}.html")
            try:
                with open(error_html_path, "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                logging.info(f"HTML salvo em: {error_html_path}")
            except Exception as save_error:
                logging.error(f"Erro ao tentar salvar HTML de depuração: {save_error}")
            raise

    except Exception as e:
        logging.error(f"Erro ao interagir com o menu de contexto: {e}")
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        error_html_path = os.path.join(DEBUG_DIR, f"error_html_context_menu_error_{timestamp}.html")
        try:
            with open(error_html_path, "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logging.info(f"HTML da página salvo em: {error_html_path}")
        except Exception as save_error:
            logging.error(f"Erro ao tentar salvar HTML de depuração: {save_error}")
        raise


### **Código Principal para Execução**

def run_automation():
    """
    Função principal que orquestra a automação do navegador.
    """
    # Caminho para o diretório de downloads
    download_dir = os.path.join(os.getcwd(), "downloads_automatizados")
    os.makedirs(download_dir, exist_ok=True)
    logging.info(f"Configurando diretório de downloads para: {download_dir}")

    # Configurações do Edge
    edge_options = EdgeOptions() 
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("--disable-popup-blocking")
    edge_options.add_argument("--disable-notifications")
    # Atenção: Se o perfil do usuário for um caminho dinâmico ou mudar,
    # considere remover ou ajustar esta linha:
    edge_options.add_argument(
        r"user-data-dir=C:\Users\2160036544\Downloads\Demandas\teste\gedocReindexações"
    )
    edge_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False, 
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    })

    service = EdgeService(r"C:\Users\2160036544\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        logging.info("1) Abrindo site principal…")
        driver.get("https://im40333.onbaseonline.com/203WPT/index.html?application=RH&locale=pt-BR")

        injetar_override_window_open(driver) 

        logging.info("2) Aguardando iframe 'htmlContent' e trocando o foco para ele…")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "htmlContent"))
        )
        logging.info("Foco mudado para o iframe 'htmlContent'.")

        injetar_override_window_open(driver) 

        logging.info("3) Inserindo matrícula…")
        matricula_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/form/div/ul/li[2]/div[2]/input"))
        )
        matricula_input.clear()
        matricula_input.send_keys("2469189")

        # 4) Clica no botão Buscar e captura a URL da popup
        clicar_no_botao(driver)
        url_popup_buscar = capturar_url_popup(driver)
        if not url_popup_buscar:
            logging.error("Não foi possível capturar a URL da popup do botão 'Buscar'.")
            return
        
        abrir_url_manual(driver, url_popup_buscar)
        
        injetar_override_window_open(driver)

        # --- PASSO CRÍTICO 1: Mudar para o iframe 'DocSelectPage' na nova aba ---
        logging.info("Mudando para o iframe 'DocSelectPage' na nova aba para acessar o conteúdo.")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "DocSelectPage"))
        )
        logging.info("Foco mudado para o iframe 'DocSelectPage'.")

        injetar_override_window_open(driver)

        # --- PASSO CRÍTICO 2: Mudar para o iframe 'frameDocSelect' DENTRO do iframe 'DocSelectPage' ---
        logging.info("Mudando para o iframe 'frameDocSelect' dentro do iframe 'DocSelectPage' para acessar a tabela.")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "frameDocSelect"))
        )
        logging.info("Foco mudado para o iframe 'frameDocSelect'.")
        
        injetar_override_window_open(driver) 

        # 6) Processa a seleção e desselecionar linhas
        selecionar_e_desselecionar_linhas(driver)

        # --- NOVO PASSO: Clicar com o botão direito e navegar pelo menu de contexto ---
        # AGORA, essa função chama a nova lógica otimizada
        clicou_menu_contexto = interagir_menu_contexto_salvar_arquivo(driver)
        # --- FIM NOVO PASSO ---
        
        # --- NOVO PASSO: Interagir com o modal 'Salvar em Arquivo' ---
        # Só tenta interagir com o modal se o clique no menu de contexto foi bem-sucedido
        if clicou_menu_contexto:
            interagir_com_modal_salvar_arquivo(driver) 
        else:
            logging.error("O clique no menu de contexto ou a abertura do modal 'Arquivo' falhou. Não é possível prosseguir com a interação do modal.")
        # --- FIM NOVO PASSO ---

        logging.info("Processo de automação concluído. Pressione Enter para fechar o navegador.")
        input() # Mantém o navegador aberto até você pressionar Enter

    except Exception as e:
        logging.error(f"Erro durante a execução principal: {e}")

    finally:
        driver.quit() # Garante que o navegador seja fechado mesmo se ocorrer um erro


# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    run_automation()