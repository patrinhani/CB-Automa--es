import subprocess
import sys
import os
import re

# --- Funções Utilitárias ---

def executar_comando(comando, descricao, suprimir_saida=False):
    """Executa um comando no terminal, com mensagens de status claras."""
    try:
        print(f"▶️  Executando: {' '.join(comando)}")
        # Opção para esconder a saída de comandos menos importantes
        stdout_pipe = subprocess.DEVNULL if suprimir_saida else None
        stderr_pipe = subprocess.DEVNULL if suprimir_saida else None
        
        subprocess.check_call(comando, stdout=stdout_pipe, stderr=stderr_pipe)
        print(f"✅  {descricao}\n")
        return True
    except subprocess.CalledProcessError as e:
        # Se falhar, tenta rodar de novo sem suprimir a saída para o usuário ver o erro
        print(f"⚠️  Comando falhou. Tentando novamente com saída detalhada...")
        try:
            subprocess.check_call(comando)
        except subprocess.CalledProcessError as detailed_e:
            print(f"❌  ERRO DETALHADO ao executar '{' '.join(comando)}'. Código de erro: {detailed_e.returncode}")
        return False
    except FileNotFoundError:
        print(f"❌  ERRO: O comando '{comando[0]}' não foi encontrado.")
        return False

def garantir_pipreqs():
    """Verifica se o pipreqs está instalado e, se não, o instala."""
    print("[PASSO 1/5] Verificando a ferramenta 'pipreqs'...")
    try:
        executar_comando([sys.executable, "-m", "pipreqs.pipreqs", "--version"], "Pipreqs já está instalado.", suprimir_saida=True)
    except subprocess.CalledProcessError:
        print("🟡  'pipreqs' não encontrado. Instalando...")
        if not executar_comando([sys.executable, "-m", "pip", "install", "pipreqs"], "'pipreqs' instalado."):
            return False
    return True

# --- Funções do Processo Principal ---

def gerar_lista_base_de_pacotes(temp_file="requirements.tmp"):
    """Usa pipreqs para gerar uma lista de pacotes base."""
    print(f"[PASSO 2/5] Detectando bibliotecas nos arquivos .py...")
    comando = ["pipreqs", ".", "--force", "--savepath", temp_file, "--encoding=utf-8"]
    return executar_comando(comando, f"Lista de pacotes base gerada em '{temp_file}'.")

def extrair_nomes_dos_pacotes(temp_file="requirements.tmp"):
    """Lê o arquivo temporário e extrai apenas os nomes dos pacotes, sem versões."""
    print(f"[PASSO 3/5] Extraindo nomes dos pacotes para resolução...")
    if not os.path.exists(temp_file):
        print(f"❌ ERRO: Arquivo temporário '{temp_file}' não encontrado.")
        return None
    
    nomes_pacotes = []
    with open(temp_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#'):
                # Usa regex para pegar apenas o nome do pacote antes de qualquer comparador
                match = re.match(r'^[a-zA-Z0-9_.-]+', line)
                if match:
                    nomes_pacotes.append(match.group(0))
    
    os.remove(temp_file) # Limpa o arquivo temporário
    print(f"✅  Nomes de pacotes extraídos: {', '.join(nomes_pacotes)}\n")
    return nomes_pacotes

def instalar_com_resolucao(nomes_pacotes):
    """Instala pacotes a partir de uma lista de nomes, deixando o pip resolver as versões."""
    print(f"[PASSO 4/5] Instalando pacotes com resolução automática de conflitos...")
    if not nomes_pacotes:
        print("🟡  Nenhum pacote para instalar.")
        return True
    
    comando = [sys.executable, "-m", "pip", "install"] + nomes_pacotes
    return executar_comando(comando, "Dependências instaladas com sucesso.")

def congelar_ambiente(lock_file="requirements_lock.txt"):
    """Salva as versões exatas do ambiente funcional em um arquivo de 'lock'."""
    print(f"[PASSO 5/5] Salvando ambiente funcional em '{lock_file}'...")
    with open(lock_file, 'w', encoding='utf-8') as f:
        comando = [sys.executable, "-m", "pip", "freeze"]
        # Redireciona a saída do comando para o arquivo
        subprocess.call(comando, stdout=f)
    print(f"✅  Ambiente 'travado' com sucesso em '{lock_file}'.\n")

def executar_passos_pos_instalacao(lock_file="requirements_lock.txt"):
    """Executa ações adicionais para pacotes específicos."""
    if not os.path.exists(lock_file): return

    with open(lock_file, 'r', encoding='utf-8') as f:
        conteudo_lock = f.read().lower()

    if "playwright" in conteudo_lock:
        executar_comando([sys.executable, "-m", "playwright", "install"], "Navegadores do Playwright baixados.")
    
    if "pytesseract" in conteudo_lock:
        print("\n" + "="*50)
        print("⚠️  AVISO MUITO IMPORTANTE SOBRE O PYTESSERACT ⚠️")
        print("Lembre-se que 'pytesseract' precisa do programa Tesseract-OCR instalado no seu computador.")
        print("Este script NÃO instala o programa. Baixe-o em:")
        print("https://github.com/UB-Mannheim/tesseract/wiki")
        print("="*50)


# --- FLUXO DE EXECUÇÃO ---

if __name__ == "__main__":
    print("======================================================")
    print("  INSTALADOR 100% AUTOMÁTICO DE AMBIENTE PYTHON       ")
    print("======================================================")

    if garantir_pipreqs():
        if gerar_lista_base_de_pacotes():
            nomes_pacotes = extrair_nomes_dos_pacotes()
            if nomes_pacotes is not None and instalar_com_resolucao(nomes_pacotes):
                congelar_ambiente()
                executar_passos_pos_instalacao()
                print("\n🎉🎉🎉 SUCESSO! O ambiente está pronto para uso. 🎉🎉🎉")
            else:
                print("\n🚨 FALHA: Não foi possível instalar as dependências.")
        else:
            print("\n🚨 FALHA: Não foi possível gerar a lista de dependências.")
    
    input("\nPressione Enter para sair.")