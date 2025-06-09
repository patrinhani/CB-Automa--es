from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support.ui import Select
import logging
import time
import os
import glob

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

# Definir um diretório para salvar logs de depuração
DEBUG_DIR = "selenium_debug_logs"
os.makedirs(DEBUG_DIR, exist_ok=True)


def save_page_html(driver, step_name):
    """Salva o HTML da página atual em um arquivo de depuração."""
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    filename = os.path.join(DEBUG_DIR, f"html_step_{step_name}_{timestamp}.html")
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        logging.info(f"HTML da página salvo para '{step_name}' em: {filename}")
    except Exception as e:
        logging.error(f"Erro ao salvar HTML para '{step_name}': {e}")


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
        save_page_html(driver, "popup_capture_timeout") # Salva HTML em caso de timeout
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
    save_page_html(driver, "after_buscar_button_click")


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
    save_page_html(driver, "after_manual_url_open")


def selecionar_e_desselecionar_linhas(driver):
    """
    Percorre todas as linhas da tabela. Tenta selecionar as linhas "RH - ANEXOS"
    simulando um clique JavaScript. Desseleciona qualquer outra linha.
    Esta função assume que o foco do Selenium JÁ ESTÁ no iframe correto.
    """
    logging.info("Iniciando seleção e desseleção de linhas na tabela...")

    try:
        # Espera que a tabela esteja visível e que as linhas com data-id existam
        WebDriverWait(driver, 30).until( # Aumentado o tempo de espera aqui
            EC.visibility_of_element_located((By.XPATH, "//tr[@role='row'][@data-id]"))
        )
        logging.info("Pelo menos uma linha da tabela está visível. Aguardando um momento para carregar todas.")
        time.sleep(5) # Aumentado a pausa para garantir que todos os dados sejam renderizados

        # Rola a tabela para garantir que todos os elementos sejam carregados no DOM
        logging.info("Tentando rolar a tabela para garantir que todas as linhas sejam carregadas.")
        grid_scroll_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "primaryHitlist_grid_scroll"))
        )
        last_height = driver.execute_script("return arguments[0].scrollHeight", grid_scroll_element)
        while True:
            driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight);", grid_scroll_element)
            time.sleep(2) # Pausa para o scroll e carregamento
            new_height = driver.execute_script("return arguments[0].scrollHeight", grid_scroll_element)
            if new_height == last_height:
                break
            last_height = new_height
        logging.info("Tabela rolada até o final. Todas as linhas devem estar disponíveis.")
        time.sleep(2) # Pausa adicional após a rolagem completa
        save_page_html(driver, "after_table_scroll")


        todas_as_linhas = driver.find_elements(By.XPATH, "//tr[@role='row'][@data-id]")
        logging.info(f"Encontradas {len(todas_as_linhas)} linhas na tabela para processar.")

        if not todas_as_linhas:
            logging.warning("Nenhuma linha encontrada na tabela. Verifique se os dados carregaram ou o XPath para as linhas.")
            save_page_html(driver, "no_table_rows_found")
            return

        # Lista para armazenar as linhas "RH - ANEXOS" que devem ser selecionadas
        linhas_rh_anexos_para_selecionar = []

        # Primeiro, identificar as linhas e desselecionar as não-RH-ANEXOS
        for i, row in enumerate(todas_as_linhas):
            try:
                doc_type_cell = row.find_element(
                    By.XPATH,
                    "./td[@aria-describedby='primaryHitlist_grid_DocumentTypeName']"
                )
                doc_type_text = doc_type_cell.text.strip()
                row_data_id = row.get_attribute('data-id')

                is_selected = row.get_attribute('aria-selected') == 'true'

                if doc_type_text == "RH - ANEXOS":
                    linhas_rh_anexos_para_selecionar.append(row)
                    if not is_selected:
                        logging.info(f"Identificada linha 'RH - ANEXOS' (data-id: {row_data_id}) que precisa ser selecionada.")
                else:
                    # Desseleciona explicitamente outras linhas se estiverem selecionadas
                    if is_selected:
                        logging.info(f"Desselecionando linha {i+1} (data-id: {row_data_id}, Tipo: '{doc_type_text}') pois não é 'RH - ANEXOS'.")
                        # Tenta um clique para desselecionar
                        driver.execute_script("arguments[0].click();", row)
                        time.sleep(0.1) # Pequena pausa para o clique processar
                        # Garante que os atributos estejam corretos caso o clique não funcione
                        driver.execute_script("arguments[0].setAttribute('aria-selected', 'false'); arguments[0].classList.remove('ui-iggrid-selectedrow');", row)
                        cells_in_row = row.find_elements(By.TAG_NAME, "td")
                        for cell in cells_in_row:
                            driver.execute_script("arguments[0].classList.remove('ui-iggrid-selectedcell', 'ui-state-active');", cell)
                    else:
                        logging.info(f"Linha {i+1} (data-id: {row_data_id}, Tipo: '{doc_type_text}') já estava desselecionada.")

            except StaleElementReferenceException:
                logging.warning(f"StaleElementReferenceException ao pré-processar a linha {i+1}. Re-tentando na próxima iteração.")
                continue
            except Exception as row_process_error:
                logging.error(f"Falha ao pré-processar a linha {i+1} (data-id: {row.get_attribute('data-id') if row else 'N/A'}): {row_process_error}")
                save_page_html(driver, f"row_preprocess_error_item_{i+1}")
                raise

        logging.info(f"Total de linhas 'RH - ANEXOS' para selecionar: {len(linhas_rh_anexos_para_selecionar)}")

        # Agora, selecionar as linhas "RH - ANEXOS" usando CTRL + Clique para seleção múltipla
        if linhas_rh_anexos_para_selecionar:
            actions = ActionChains(driver)
            actions.key_down(Keys.CONTROL) # Pressiona CTRL

            for i, row in enumerate(linhas_rh_anexos_para_selecionar):
                try:
                    row_data_id = row.get_attribute('data-id')
                    is_selected = row.get_attribute('aria-selected') == 'true'
                    if not is_selected:
                        actions.click(row) # Clica na linha com CTRL pressionado
                        logging.info(f"Clicando com CTRL na linha 'RH - ANEXOS' (data-id: {row_data_id}) para seleção.")
                        actions.perform() # Executa a ação
                        # IMPORTANTE: Reiniciar ActionChains para garantir que CTRL continue pressionado
                        actions = ActionChains(driver)
                        actions.key_down(Keys.CONTROL)
                        time.sleep(0.2) # Pequena pausa
                        # Verificação imediata após o clique (pode ser útil para depuração)
                        if row.get_attribute('aria-selected') != 'true':
                            logging.warning(f"Atenção: Linha 'RH - ANEXOS' (data-id: {row_data_id}) não parece ter sido selecionada via CTRL+Click. Tentando forçar via JS.")
                            driver.execute_script("arguments[0].setAttribute('aria-selected', 'true'); arguments[0].classList.add('ui-iggrid-selectedrow');", row)
                            cells_in_row = row.find_elements(By.TAG_NAME, "td")
                            for cell in cells_in_row:
                                driver.execute_script("arguments[0].classList.add('ui-iggrid-selectedcell', 'ui-state-active');", cell)
                            logging.info(f"Forçado seleção de Linha 'RH - ANEXOS' (data-id: {row_data_id}) via setAttribute/classList.")
                    else:
                        logging.info(f"Linha 'RH - ANEXOS' (data-id: {row_data_id}) já estava selecionada.")
                except StaleElementReferenceException:
                    logging.warning(f"StaleElementReferenceException ao tentar selecionar a linha {i+1} (data-id: {row.get_attribute('data-id') if row else 'N/A'}). Pode ser que o DOM foi re-renderizado. Continuando...")
                    continue
                except Exception as click_error:
                    logging.error(f"Erro ao tentar CTRL+Click na linha 'RH - ANEXOS' (data-id: {row.get_attribute('data-id') if row else 'N/A'}): {click_error}")
                    save_page_html(driver, f"ctrl_click_error_item_{i+1}")
                    # Tenta forçar a seleção via JS como fallback
                    driver.execute_script("arguments[0].setAttribute('aria-selected', 'true'); arguments[0].classList.add('ui-iggrid-selectedrow');", row)
                    cells_in_row = row.find_elements(By.TAG_NAME, "td")
                    for cell in cells_in_row:
                        driver.execute_script("arguments[0].classList.add('ui-iggrid-selectedcell', 'ui-state-active');", cell)
                    logging.info(f"Forçado seleção de Linha 'RH - ANEXOS' (data-id: {row.get_attribute('data-id') if row else 'N/A'}) via setAttribute/classList como fallback.")

            actions.key_up(Keys.CONTROL).perform() # Libera CTRL no final
            logging.info("Seleção de linhas 'RH - ANEXOS' com CTRL + Clique concluída.")
        else:
            logging.warning("Nenhuma linha 'RH - ANEXOS' encontrada para selecionar.")

        logging.info("Processo de seleção/desseleção de linhas concluído.")
        save_page_html(driver, "after_selection_process")


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
        save_page_html(driver, "global_selection_error")
        raise


def interagir_com_modal_salvar_arquivo(driver, download_dir):
    """
    Interage com o modal/janela que aparece após clicar em 'Arquivo' no menu de contexto.
    Isso inclui selecionar 'PDF' e clicar em 'Salvar'.
    """
    logging.info("Iniciando interação com o modal 'Salvar em Arquivo'...")
    try:
        # Muda o foco para o conteúdo padrão da janela, onde o modal geralmente aparece
        driver.switch_to.default_content()
        logging.info("Foco do Selenium mudado para o conteúdo padrão da janela para interagir com o modal.")
        save_page_html(driver, "before_modal_interaction")


        # Espera que o modal "Salvar em Arquivo" esteja visível.
        WebDriverWait(driver, 20).until( # Aumentado o tempo de espera
            EC.visibility_of_element_located((By.ID, "frmSaveToFile"))
        )
        logging.info("Modal 'Salvar em Arquivo' detectado e visível.")
        save_page_html(driver, "modal_visible")
        time.sleep(2) # Pausa para garantir que todos os elementos internos do modal estejam carregados

        logging.info("Selecionando 'PDF (.pdf)' no dropdown 'Salvar como'...")
        select_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "selectContent"))
        )
        select_obj = Select(select_element)
        select_obj.select_by_value("pdf")
        logging.info("Opção 'PDF (.pdf)' selecionada.")
        time.sleep(1) # Pequena pausa

        logging.info("Clicando no botão 'Salvar'...")
        save_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnSave"))
        )
        save_button.click()
        logging.info("Botão 'Salvar' clicado.")
        save_page_html(driver, "after_modal_save_click")
        time.sleep(5) # Uma pausa maior para o download começar a processar

        try:
            WebDriverWait(driver, 15).until(
                EC.invisibility_of_element_located((By.ID, "frmSaveToFile"))
            )
            logging.info("Modal 'Salvar em Arquivo' desapareceu.")
        except TimeoutException:
            logging.warning("Modal 'Salvar em Arquivo' não desapareceu em 15 segundos, pode estar processando ou necessitar de outra ação.")
            save_page_html(driver, "modal_not_disappeared_timeout")

        # Verificação do download
        logging.info("Verificando se o(s) arquivo(s) PDF foi(foram) baixado(s)...")
        try:
            end_time = time.time() + 45 # Espera até 45 segundos pelo download (pode ser um processo demorado)
            downloaded_success = False
            while time.time() < end_time:
                pdf_files = glob.glob(os.path.join(download_dir, "*.pdf"))
                # Filtra apenas por arquivos PDF que não estão sendo baixados (.crdownload, .tmp, etc.)
                downloaded_pdf_files = [f for f in pdf_files if not f.endswith(('.crdownload', '.tmp', '.part'))]
                if downloaded_pdf_files:
                    logging.info(f"Arquivo(s) PDF baixado(s) com sucesso: {downloaded_pdf_files}")
                    downloaded_success = True
                    break
                time.sleep(2) # Verifica a cada 2 segundos
            if not downloaded_success:
                logging.warning("Timeout: Nenhum arquivo PDF encontrado no diretório de download após o clique em salvar.")
                save_page_html(driver, "download_not_found_timeout")

        except Exception as download_check_error:
            logging.error(f"Erro ao verificar download: {download_check_error}")
            save_page_html(driver, "download_check_error")


    except Exception as e:
        logging.error(f"Erro ao interagir com o modal 'Salvar em Arquivo': {e}")
        save_page_html(driver, "modal_interaction_error")
        raise
    finally:
        # Retorna o foco para o iframe principal após a interação com o modal
        driver.switch_to.default_content()
        try:
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "DocSelectPage")))
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameDocSelect")))
            logging.info("Foco do Selenium retornado para o iframe 'frameDocSelect' após interagir com o modal.")
        except TimeoutException:
            logging.warning("Não foi possível retornar o foco para os iframes principais. A página pode ter navegado ou o DOM mudado significativamente.")
            save_page_html(driver, "iframe_return_focus_timeout")


def interagir_menu_contexto_salvar_arquivo(driver):
    """
    Executa o clique com o botão direito, simula hover e chama o handler JS
    para a opção 'Arquivo' do menu de contexto.
    """
    logging.info("Iniciando interação com o menu de contexto (Solução Otimizada)...")
    original_all_handles = driver.window_handles

    try:
        # 1. Encontra a primeira célula "RH - ANEXOS" selecionada para o clique direito
        try:
            first_selected_rh_anexos_cell = WebDriverWait(driver, 15).until( # Aumentado o tempo de espera
                EC.visibility_of_element_located((
                    By.XPATH,
                    "//tr[@role='row'][@data-id][@aria-selected='true']"
                    "/td[@aria-describedby='primaryHitlist_grid_DocumentTypeName'"
                    " and normalize-space(text())='RH - ANEXOS']"
                ))
            )
            logging.info(f"Primeira célula 'RH - ANEXOS' selecionada encontrada para clique direito: {first_selected_rh_anexos_cell.text}")
        except TimeoutException:
            logging.error("Timeout: Nenhuma linha 'RH - ANEXOS' visível e selecionada para o clique direito.")
            save_page_html(driver, "no_selected_row_for_right_click")
            return False


        # 2. Dispara o evento 'contextmenu' via JavaScript na *célula da linha selecionada*.
        driver.execute_script("arguments[0].dispatchEvent(new Event('contextmenu', { bubbles: true }));", first_selected_rh_anexos_cell)
        logging.info("Evento 'contextmenu' disparado via JavaScript na linha selecionada.")
        time.sleep(1) # Pequena pausa para o menu de contexto aparecer
        save_page_html(driver, "after_context_menu_event")


        # 3. Espera pelo item "Enviar para" (ID: menuControl_25) no menu de contexto principal.
        enviar_para_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "menuControl_25"))
        )
        logging.info("Item 'Enviar para' do menu de contexto (ID: menuControl_25) visível e clicável.")

        # 4. Simula o hover em "Enviar para" disparando um evento mouseover via JavaScript.
        driver.execute_script("arguments[0].dispatchEvent(new Event('mouseover', { bubbles: true }));", enviar_para_item)
        logging.info("Evento mouseover disparado em 'Enviar para' via JavaScript (para hover).")

        time.sleep(2.5) # PAUSA MAIOR para garantir que o submenu "Arquivo" apareça e estabilize. CRÍTICO.
        save_page_html(driver, "after_send_to_hover")


        # 5. Tenta encontrar o item "Arquivo" (ID: menuControl_13) no submenu e clica nele
        arquivo_item_clicked = False
        try:
            # Tenta encontrar e clicar no elemento visível do submenu
            arquivo_item = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "menuControl_13"))
            )
            logging.info("Item 'Arquivo' do submenu (ID: menuControl_13) visível e clicável. Clicando...")
            arquivo_item.click() # Tenta o clique direto no elemento
            logging.info("Clique nativo no item 'Arquivo' executado.")
            arquivo_item_clicked = True
        except TimeoutException:
            logging.warning("Item 'Arquivo' (ID: menuControl_13) não encontrado para clique nativo ou não clicável. Tentando chamar JS handler como alternativa.")

        if not arquivo_item_clicked:
            # Fallback: Chama diretamente o handler JavaScript para o item 'Arquivo' usando seu ID ('13').
            try:
                WebDriverWait(driver, 10).until(
                    lambda dr: dr.execute_script("return window.contextMenu && typeof window.contextMenu.onMenuItemSelected === 'function';")
                )
                logging.info("Objeto JavaScript 'window.contextMenu' e método 'onMenuItemSelected' detectados.")
                logging.info("Chamando diretamente o handler JavaScript 'onMenuItemSelected' para o item 'Arquivo' (ID 13).")
                driver.execute_script("window.contextMenu.onMenuItemSelected('13');")
                logging.info("Handler para 'Arquivo' executado via JavaScript.")
                arquivo_item_clicked = True
            except TimeoutException:
                logging.error("Timeout: Objeto 'window.contextMenu' ou método 'onMenuItemSelected' não disponível.")
                save_page_html(driver, "js_handler_not_available_timeout")
                return False
            except Exception as js_handler_error:
                logging.error(f"Erro ao chamar JS handler para 'Arquivo': {js_handler_error}")
                save_page_html(driver, "js_handler_call_error")
                return False

        if not arquivo_item_clicked:
            logging.error("Falha ao clicar ou ativar o item 'Arquivo' do menu de contexto por qualquer método.")
            save_page_html(driver, "file_menu_item_activation_failed")
            return False

        # Limpa qualquer URL capturada anteriormente antes de verificar por nova janela.
        driver.execute_script("window.openedURL = null;")
        logging.info("window.openedURL resetado antes de verificar nova janela.")
        time.sleep(1) # Pequena pausa para a ação iniciar e o possível popup/modal ser criado
        save_page_html(driver, "after_file_menu_item_activation")


        # 6. Verifica se uma nova janela/guia foi aberta (onde o modal de salvar pode estar)
        # OU se o modal apareceu na mesma aba.
        logging.info("Verificando se uma nova janela/guia foi aberta para o modal de download...")
        try:
            # Espera que o número de janelas aumente, indicando uma nova janela/tab do modal
            WebDriverWait(driver, 15).until(EC.number_of_windows_to_be(len(original_all_handles) + 1))
            new_window_handle = [handle for handle in driver.window_handles if handle not in original_all_handles][0]
            driver.switch_to.window(new_window_handle)
            logging.info(f"Nova janela/guia detectada para o modal de download com URL: {driver.current_url}")
            save_page_html(driver, "new_window_detected")
            return True # Indica sucesso na abertura de nova janela
        except TimeoutException:
            # Se não abriu nova janela, o modal pode ter aparecido na mesma aba.
            logging.warning("Timeout: Nenhuma nova janela/guia foi detectada em 15 segundos após a execução do handler 'Arquivo'. Assumindo que o modal está na mesma aba ou é um overlay.")
            save_page_html(driver, "no_new_window_timeout")
            return True # Indica que a ação do menu de contexto foi, pelo menos, executada.
        except Exception as win_handle_error:
            logging.error(f"Erro ao verificar nova janela/guia após clique no menu: {win_handle_error}")
            save_page_html(driver, "new_window_check_error")
            return False

    except Exception as e:
        logging.error(f"Erro geral ao interagir com o menu de contexto: {e}")
        save_page_html(driver, "context_menu_global_error")
        raise


def run_automation():
    """
    Função principal que orquestra a automação do navegador.
    """
    download_dir = os.path.join(os.getcwd(), "downloads_automatizados")
    os.makedirs(download_dir, exist_ok=True)
    logging.info(f"Configurando diretório de downloads para: {download_dir}")

    edge_options = EdgeOptions()
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("--disable-popup-blocking")
    edge_options.add_argument("--disable-notifications")
    edge_options.add_argument(
        r"user-data-dir=C:\Users\2160036544\Downloads\Demandas\teste\gedocReindexações"
    )
    edge_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    })

    edgedriver_path = r"C:\Users\2160036544\Downloads\edgedriver_win64\msedgedriver.exe"
    if not os.path.exists(edgedriver_path):
        logging.error(f"msedgedriver.exe não encontrado em: {edgedriver_path}. Por favor, verifique o caminho.")
        return

    service = EdgeService(edgedriver_path)
    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        logging.info("1) Abrindo site principal…")
        driver.get("https://im40333.onbaseonline.com/203WPT/index.html?application=RH&locale=pt-BR")
        save_page_html(driver, "initial_load")

        injetar_override_window_open(driver)

        logging.info("2) Aguardando iframe 'htmlContent' e trocando o foco para ele…")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "htmlContent"))
        )
        logging.info("Foco mudado para o iframe 'htmlContent'.")
        save_page_html(driver, "switched_to_htmlContent_iframe")


        injetar_override_window_open(driver)

        logging.info("3) Inserindo matrícula…")
        matricula_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/form/div/ul/li[2]/div[2]/input"))
        )
        matricula_input.clear()
        matricula_input.send_keys("2469189")

        clicar_no_botao(driver)
        url_popup_buscar = capturar_url_popup(driver)
        if not url_popup_buscar:
            logging.error("Não foi possível capturar a URL da popup do botão 'Buscar'.")
            return

        abrir_url_manual(driver, url_popup_buscar)

        injetar_override_window_open(driver)

        logging.info("Mudando para o iframe 'DocSelectPage' na nova aba para acessar o conteúdo.")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "DocSelectPage"))
        )
        logging.info("Foco mudado para o iframe 'DocSelectPage'.")
        save_page_html(driver, "switched_to_docselectpage_iframe")


        injetar_override_window_open(driver)

        logging.info("Mudando para o iframe 'frameDocSelect' dentro do iframe 'DocSelectPage' para acessar a tabela.")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "frameDocSelect"))
        )
        logging.info("Foco mudado para o iframe 'frameDocSelect'.")
        save_page_html(driver, "switched_to_framedocselect_iframe")


        injetar_override_window_open(driver)

        selecionar_e_desselecionar_linhas(driver)

        clicou_menu_contexto = interagir_menu_contexto_salvar_arquivo(driver)

        if clicou_menu_contexto:
            interagir_com_modal_salvar_arquivo(driver, download_dir)
        else:
            logging.error("O clique no menu de contexto ou a abertura do modal 'Arquivo' falhou. Não é possível prosseguir com a interação do modal.")

        logging.info("Processo de automação concluído. Pressione Enter para fechar o navegador.")
        input()

    except Exception as e:
        logging.error(f"Erro durante a execução principal: {e}")
        save_page_html(driver, "main_execution_error")

    finally:
        driver.quit()


if __name__ == "__main__":
    run_automation()