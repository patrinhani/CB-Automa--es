import time
import keyboard
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options # Importa Options para configurar o Edge

# --- Configuração do EdgeDriver com Caminho Específico ---
caminho_para_edgedriver = r"C:\Users\2160036544\Downloads\edgedriver_win64\msedgedriver.exe"
service = EdgeService(executable_path=caminho_para_edgedriver)

# --- Configuração de Opções para o Edge ---
edge_options = Options()
# Adiciona argumento para ignorar erros de certificado
edge_options.add_argument("--ignore-certificate-errors")
# Adiciona argumento para permitir conteúdo não seguro (misturado) - pode ser útil
edge_options.add_argument("--allow-running-insecure-content")
# Tenta usar a opção para desativar a verificação de segurança QUIC (raramente necessária, mas vale a pena tentar)
# edge_options.add_argument("--disable-quic")
# edge_options.add_argument("--disable-ssl-false-start") # Outra opção menos comum

print(f"Usando EdgeDriver do caminho: {caminho_para_edgedriver}")
print("Tentando ignorar erros de certificado e permitir conteúdo inseguro.")

# --- Abrir o Navegador Edge e Acessar o Link ---
url_do_site = "https://documentos/bahia/gateway" # SEU LINK INTERNO/LOCAL

driver = None

try:
    # Inicializar o navegador Edge com o serviço e as opções configuradas
    driver = webdriver.Edge(service=service, options=edge_options)
    print(f"Abrindo o navegador Edge e acessando: {url_do_site}")

    driver.get(url_do_site)

    print(f"Página '{url_do_site}' acessada com sucesso (ou com erros de certificado ignorados)!")
    print("\nO navegador permanecerá aberto. Pressione 'ESC' para fechar.")

    # Loop para manter o navegador aberto até ESC ser pressionado
    while True:
        if keyboard.is_pressed('esc'):
            print("Tecla 'ESC' detectada. Fechando o navegador...")
            break
        time.sleep(0.1)

except Exception as e:
    print(f"Ocorreu um erro: {e}")
    # Se a página não carregou ou deu erro, o erro pode ser no `driver.get()`
    # O navegador pode exibir a página de erro SSL.

finally:
    if driver:
        driver.quit()
        print("Navegador fechado.")