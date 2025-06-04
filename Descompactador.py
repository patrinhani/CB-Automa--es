import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
import stat # Para manter permissões de arquivo
import shutil

# --- Configuração de Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Função Auxiliar para Caminhos Longos no Windows ---
# Esta função adiciona o prefixo '\\?\' para contornar o limite MAX_PATH no Windows
# para caminhos com mais de 260 caracteres.
def _winapi_path(path):
    if os.name == 'nt': # Se for Windows
        # Converte para caminho absoluto e depois adiciona o prefixo
        path = os.path.abspath(path)
        if path.startswith("\\\\"): # UNC path (rede)
            return "\\\\?\\UNC\\" + path[2:]
        else:
            return "\\\\?\\" + path
    return path # Retorna o caminho original para outros SOs ou se já estiver ok

# --- Função Principal de Descompactação ---
def descompactar_com_correcao_caminho_longo(zip_path, extract_base_dir):
    """
    Descompacta um arquivo ZIP, lidando com problemas de caminho muito longo no Windows.
    Os arquivos serão descompactados em uma subpasta com o nome do ZIP dentro de extract_base_dir.

    Args:
        zip_path (str): Caminho completo para o arquivo ZIP.
        extract_base_dir (str): Diretório base onde a nova subpasta será criada.
    """
    if not os.path.exists(zip_path):
        messagebox.showerror("Erro", f"O arquivo ZIP não foi encontrado: {zip_path}")
        logging.error(f"Arquivo ZIP não encontrado: {zip_path}")
        return

    # --- NOVIDADE AQUI: Define o diretório de destino específico ---
    # Obtém o nome do arquivo ZIP sem a extensão
    zip_name_without_ext = os.path.splitext(os.path.basename(zip_path))[0]
    # Cria o caminho completo para a nova subpasta de destino
    extract_dir = os.path.join(extract_base_dir, zip_name_without_ext)
    # --- FIM DA NOVIDADE ---

    # Garante que o diretório de destino existe e tem o prefixo para caminhos longos
    extract_dir_long_path = _winapi_path(extract_dir)
    os.makedirs(extract_dir_long_path, exist_ok=True)
    logging.info(f"Diretório de destino para descompactação: {extract_dir_long_path}")

    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            for member in zip_ref.infolist():
                # Caminho completo do membro dentro do ZIP no destino FINAL
                member_full_path = os.path.join(extract_dir, member.filename)
                
                # Aplica a correção de caminho longo antes de tentar extrair
                target_path = _winapi_path(member_full_path)
                
                # Se for um diretório, cria-o e continua para o próximo membro
                if member.is_dir():
                    os.makedirs(target_path, exist_ok=True)
                    logging.debug(f"Criando diretório: {target_path}")
                    continue

                # Para arquivos, extraia o membro
                try:
                    # Extrair o membro. O 'path' deve ser o diretório base para onde o ZIP irá.
                    # Ele lidará com a estrutura interna do ZIP.
                    zip_ref.extract(member, path=extract_dir_long_path)
                    
                    # Tenta manter permissões do arquivo
                    perms = member.external_attr >> 16
                    if perms != 0:
                        try:
                            mode = stat.S_IMODE(perms)
                            os.chmod(target_path, mode)
                        except Exception as e_perm:
                            logging.warning(f"Não foi possível aplicar permissões para {target_path}: {e_perm}")
                            
                    logging.info(f"Extraído: {member.filename} para {target_path}")

                except OSError as e:
                    logging.warning(f"Erro ao extrair {member.filename} para {target_path}: {e}. Tentando estratégia alternativa...")
                    
                    temp_dir_base = os.path.join(os.path.expanduser("~"), "temp_zip_ext")
                    temp_dir_long_path = _winapi_path(temp_dir_base)
                    os.makedirs(temp_dir_long_path, exist_ok=True)

                    temp_file_extracted_path = os.path.join(temp_dir_long_path, member.filename)

                    try:
                        zip_ref.extract(member, path=temp_dir_long_path)
                        
                        # Move do local temporário para o destino final (que inclui a subpasta)
                        shutil.move(temp_file_extracted_path, target_path)
                        logging.info(f"Extraído (via temp dir) e movido: {member.filename} para {target_path}")
                    except Exception as move_e:
                        logging.error(f"Falha ao extrair/mover {member.filename} mesmo com temp dir: {move_e}")
                        messagebox.showerror("Erro Crítico de Extração",
                                             f"Falha ao extrair '{member.filename}' mesmo com estratégia de caminho curto. "
                                             "Pode ser necessário usar um descompactador externo (ex: 7-Zip).")
                        return # Aborta se um arquivo crítico não puder ser extraído

    except zipfile.BadZipFile:
        messagebox.showerror("Erro", "O arquivo selecionado não é um arquivo ZIP válido ou está corrompido.")
        logging.error(f"Arquivo não é um ZIP válido ou está corrompido: {zip_path}")
    except Exception as e:
        messagebox.showerror("Erro Inesperado", f"Ocorreu um erro inesperado: {e}")
        logging.error(f"Erro inesperado ao descompactar {zip_path}: {e}", exc_info=True)
    finally:
        # Tenta remover o diretório temporário após a descompactação
        if 'temp_dir_base' in locals() and os.path.exists(_winapi_path(temp_dir_base)):
            try:
                shutil.rmtree(_winapi_path(temp_dir_base))
                logging.info(f"Diretório temporário removido: {temp_dir_base}")
            except Exception as e:
                logging.warning(f"Não foi possível remover o diretório temporário {temp_dir_base}: {e}")

    messagebox.showinfo("Sucesso", f"Descompactação concluída para: {zip_path}\n"
                                   f"Arquivos extraídos para: {extract_dir}")
    logging.info(f"Descompactação concluída para {zip_path} em {extract_dir}.")

# --- Interface Gráfica (Tkinter) ---
def selecionar_arquivo_zip():
    root = tk.Tk()
    root.withdraw() # Oculta a janela principal do Tkinter

    zip_file_path = filedialog.askopenfilename(
        title="Selecione o arquivo ZIP para descompactar",
        filetypes=[("Arquivos ZIP", "*.zip"), ("Todos os arquivos", "*.*")]
    )
    return zip_file_path

def selecionar_pasta_destino_base(): # Renomeada para clareza
    root = tk.Tk()
    root.withdraw() # Oculta a janela principal do Tkinter

    # Sugere a pasta Downloads do usuário como base
    default_extract_base_path = os.path.join(os.path.expanduser("~"), "Downloads", "Arquivos_Descompactados")
    
    # Cria a pasta padrão se não existir
    os.makedirs(_winapi_path(default_extract_base_path), exist_ok=True)

    extract_base_folder = filedialog.askdirectory(
        title="Selecione a PASTA BASE de destino para a descompactação",
        initialdir=default_extract_base_path # Sugere a pasta base
    )
    return extract_base_folder

# --- Execução Principal ---
if __name__ == "__main__":
    logging.info("Iniciando utilitário de descompactação de ZIP com correção de caminho longo e organização em subpasta.")

    caminho_zip = selecionar_arquivo_zip()
    if not caminho_zip:
        logging.info("Nenhum arquivo ZIP selecionado. Encerrando.")
        exit()

    # Usamos o novo nome da função aqui
    caminho_destino_base = selecionar_pasta_destino_base()
    if not caminho_destino_base:
        logging.info("Nenhuma pasta base de destino selecionada. Encerrando.")
        exit()

    descompactar_com_correcao_caminho_longo(caminho_zip, caminho_destino_base)

    logging.info("Processo de descompactação concluído.")