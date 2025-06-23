import logging
import os
import time
import tkinter as tk
from tkinter import scrolledtext
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# --- INÍCIO DA SEÇÃO: Interface Gráfica para Inserir Matrículas ---

def obter_matriculas_gui():
    """
    Cria uma janela para o usuário inserir as matrículas.
    Retorna uma lista de matrículas ou None se o usuário cancelar.
    """
    matriculas_resultado = []
    
    def on_iniciar():
        """
        Pega o texto da caixa, o divide em uma lista, remove linhas vazias
        e fecha a janela.
        """
        texto_inserido = text_area.get("1.0", tk.END).strip()
        if texto_inserido:
            linhas = texto_inserido.split('\n')
            matriculas_resultado.extend([linha.strip() for linha in linhas if linha.strip()])
        root.destroy()

    root = tk.Tk()
    root.title("Inserir Matrículas para Automação (Playwright)")

    # Centralizar a janela
    window_width = 400
    window_height = 350
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    label = tk.Label(root, text="Cole as matrículas abaixo (uma por linha):", font=("Arial", 12))
    label.pack(pady=(10, 5))

    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=40, height=10, font=("Arial", 11))
    text_area.pack(pady=5, padx=10, fill="both", expand=True)
    text_area.focus()

    button_frame = tk.Frame(root)
    button_frame.pack(pady=(5, 10))

    start_button = tk.Button(button_frame, text="Iniciar Automação", command=on_iniciar, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
    start_button.pack(side=tk.LEFT, padx=10)

    cancel_button = tk.Button(button_frame, text="Cancelar", command=root.destroy, bg="#f44336", fg="white", font=("Arial", 10))
    cancel_button.pack(side=tk.LEFT, padx=10)
    
    root.mainloop()

    return matriculas_resultado if matriculas_resultado else None

# --- FIM DA SEÇÃO DA INTERFACE ---


# Configuração do logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')


def preencher_e_enviar_lote_playwright(matriculas_para_adicionar):
    """
    Função de automação com Playwright, corrigida para ser mais robusta e específica.
    """
    caminho_perfil_edge = r"C:\Users\2160036544\Downloads\Demandas\teste\gedocReindexações"
    url_site = "https://im40333.onbaseonline.com/203WPT/index.html?application=RH&locale=pt-BR"
    
    if not os.path.exists(caminho_perfil_edge):
        logging.error(f"ERRO: O diretório de perfil do usuário não foi encontrado em: {caminho_perfil_edge}")
        return

    with sync_playwright() as p:
        context = None
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=caminho_perfil_edge,
                headless=False,
                channel="msedge",
                args=["--start-maximized"],
                slow_mo=50,
                no_viewport=True
            )
            page = context.new_page()
            page.set_default_timeout(35000) 

            page.goto(url_site)
            logging.info(f"1. Acessando o site: {url_site}")

            logging.info("2. Aguardando o painel de formulários e clicando no botão '+ Lote'...")
            seletor_lote = "#formListPanel td[data-name='RH - Solicitacao de Lote']"
            page.locator(seletor_lote).click()
            logging.info("   -> Botão '+ Lote' clicado.")

            logging.info("3. Localizando iframes 'iframeHolder' e 'uf_hostframe'...")
            frame_principal = page.frame_locator("#iframeHolder")
            frame_aninhado = frame_principal.frame_locator("#uf_hostframe")
            logging.info("   -> Foco definido nos iframes.")

            logging.info(f"4. Iniciando o cadastro de {len(matriculas_para_adicionar)} matrículas...")
            
            botao_adicionar = frame_aninhado.locator('input[aria-label*="Lista de Matrículas"]')
            
            for i, matricula_atual in enumerate(matriculas_para_adicionar):
                logging.info(f"   -> Adicionando linha {i+1} para a matrícula '{matricula_atual}'")
                botao_adicionar.click()
                
                id_campo_dinamico = f"#rhmatricula12_input_{i}"
                campo_matricula = frame_aninhado.locator(id_campo_dinamico)
                
                campo_matricula.fill(matricula_atual)
                logging.info(f"       -> Campo preenchido com sucesso.")
            
            logging.info(f"   -> Todas as {len(matriculas_para_adicionar)} matrículas foram adicionadas.")

            logging.info("5. Clicando no botão 'Enviar' para finalizar o processo...")
            botao_enviar = frame_aninhado.locator("input[value='Enviar']")
            botao_enviar.click()
            logging.info("   -> Botão 'Enviar' clicado.")
            
            logging.info("--- PROCESSO CONCLUÍDO: Matrículas enviadas com sucesso! ---")
            logging.info("O navegador será fechado em 5 segundos...")
            time.sleep(5)

        except PlaywrightTimeoutError:
            logging.error("ERRO: Um elemento (botão ou campo) não foi encontrado a tempo (timeout).")
            input("Ocorreu um erro de Timeout. Pressione Enter para fechar...") 
        except Exception as e:
            logging.error(f"Ocorreu um erro inesperado: {e}")
            input("Ocorreu um erro inesperado. Pressione Enter para fechar...")
        finally:
            if context:
                context.close()
                logging.info("Navegador fechado.")


# Bloco de execução principal
if __name__ == "__main__":
    # 1. Chamar a tela para obter as matrículas do usuário
    lista_de_matriculas = obter_matriculas_gui()

    # 2. Verificar se o usuário forneceu matrículas antes de continuar
    if lista_de_matriculas:
        logging.info(f"Matrículas recebidas: {len(lista_de_matriculas)}. Iniciando automação com Playwright...")
        # 3. Executar a automação com a lista fornecida
        preencher_e_enviar_lote_playwright(lista_de_matriculas)
    else:
        logging.warning("Nenhuma matrícula foi inserida ou a operação foi cancelada. Encerrando o programa.")