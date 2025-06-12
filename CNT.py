import os
import shutil
import fitz  # PyMuPDF
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_path
import concurrent.futures
import time

# ----------------------- CONFIGURAÇÕES -----------------------
# Configure os caminhos de acordo com a sua máquina.

# 📌 Aponta para o executável do Tesseract.
# Use o caminho da sua instalação.
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\2160036544\Downloads\tesseract-5.5.1\tesseract.exe'

# 📌 Aponta para a pasta 'bin' do Poppler.
# Use o caminho da sua instalação.
POPPLER_PATH = r'C:\Users\2160036544\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin'

# 📂 Caminho para a pasta que contém as subpastas dos funcionários.
CAMINHO_LOTE = r"C:\Users\2160036544\Downloads\SavedDocument\lote1" # <-- ALTERE PARA O CAMINHO REAL

# 📁 Pasta onde os documentos classificados serão salvos.
PASTA_ANEXO = os.path.join(CAMINHO_LOTE, "Lote - Anexo (Classificado)")
os.makedirs(PASTA_ANEXO, exist_ok=True)

# 🔑 Palavras-chave para classificação dos documentos.
PALAVRAS_CHAVE = {
    "ASO A": ["exame admissional"],
    "ASO D": ["exame demissional"],
    "ASO P": ["exame periódico"],
    "Contrato de Trabalho": ["contrato", "acordo", "compensação", "prorrogação", "bh", "docs. admissionais", "termo", "termos", "ficha de registro"],
    "Formulário de atualização de CTPS": ["ficha de anotações", "atualizações da ctps"],
    "TRCT Homologado": ["trct", "guia do seguro", "grrf", "fgts", "extrato analítico", "comprovante de crédito em conta corrente"],
    "Comunicado de Aviso de Férias": ["aviso de férias"],
    "Comunicado De Dispensa": ["telegramas", "contato sms", "e-mail", "whatsapp", "pedido", "aviso de desligamento", "término de contrato", "justa causa", "óbito", "falecimento"]
}

# ---------------------- FUNÇÕES AUXILIARES ----------------------

def extrair_texto(pdf_path):
    """
    Extrai texto de um PDF, usando OCR otimizado com OpenCV se necessário.
    """
    try:
        texto_extraido = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                texto_extraido += page.get_text("text", flags=fitz.TEXT_INHIBIT_SPACES).lower()
        
        if not texto_extraido.strip():
            # DPI reduzido para 200 para acelerar a conversão
            imagens_pil = convert_from_path(pdf_path, dpi=200, poppler_path=POPPLER_PATH)
            
            for pil_image in imagens_pil:
                # --- OTIMIZAÇÃO COM OPENCV ---
                # 1. Converte para formato OpenCV e escala de cinza
                imagem_cv = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2GRAY)
                # 2. Binariza a imagem (preto e branco) para melhorar o contraste
                _, imagem_processada = cv2.threshold(imagem_cv, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                # --- FIM DA OTIMIZAÇÃO ---
                
                texto_extraido += pytesseract.image_to_string(imagem_processada, lang='por').lower()
                
        return texto_extraido
    except Exception as e:
        print(f"❌ Erro ao extrair texto de {os.path.basename(pdf_path)}: {e}")
        return ""

def classificar_documento(matricula_nome, pdf_path, texto):
    """
    Classifica um documento e copia para a pasta de destino.
    """
    encontrados = set() # Usar um 'set' evita categorias duplicadas
    for categoria, palavras in PALAVRAS_CHAVE.items():
        if any(palavra in texto for palavra in palavras):
            encontrados.add(categoria)
    
    if encontrados:
        for tipo in encontrados:
            nome_pasta_destino = f"{matricula_nome.split('-')[0].strip()} - {tipo}"
            pasta_destino = os.path.join(PASTA_ANEXO, nome_pasta_destino)
            os.makedirs(pasta_destino, exist_ok=True)
            shutil.copy2(pdf_path, pasta_destino)
        # Imprime apenas uma vez por arquivo, com todas as suas categorias
        print(f"📎 Classificado: {os.path.basename(pdf_path)} ➤ {list(encontrados)}")
    else:
        print(f"🟡 Ignorado (sem palavra-chave): {os.path.basename(pdf_path)}")

def processar_um_arquivo(args):
    """
    Função que encapsula o trabalho para um único arquivo, para ser usada em paralelo.
    """
    pasta_matricula, arquivo, caminho_completo_pasta = args
    caminho_pdf = os.path.join(caminho_completo_pasta, arquivo)
    texto = extrair_texto(caminho_pdf)
    if texto:
        classificar_documento(pasta_matricula, caminho_pdf, texto)

# ------------------------ EXECUÇÃO PRINCIPAL ------------------------

def main():
    """
    Função principal que orquestra todo o processo de forma paralela.
    """
    inicio = time.time()
    print("🚀 Iniciando automação de classificação de prontuários (versão otimizada)...")
    
    if not os.path.exists(CAMINHO_LOTE):
        print(f"❌ ERRO: Caminho do lote não encontrado em '{CAMINHO_LOTE}'")
        return

    subpastas_matricula = [d for d in os.listdir(CAMINHO_LOTE) if os.path.isdir(os.path.join(CAMINHO_LOTE, d)) and not d.startswith("Lote - Anexo")]

    if not subpastas_matricula:
        print("⚠️ Nenhuma subpasta de matrícula encontrada para processar.")
        return

    # 1. Monta uma lista com todas as tarefas (todos os arquivos a processar)
    tarefas = []
    for pasta in subpastas_matricula:
        caminho_pasta = os.path.join(CAMINHO_LOTE, pasta)
        for arquivo in os.listdir(caminho_pasta):
            if arquivo.lower().endswith(".pdf"):
                tarefas.append((pasta, arquivo, caminho_pasta))
    
    if not tarefas:
        print("⚠️ Nenhum arquivo PDF encontrado em nenhuma das subpastas.")
        return
        
    print(f"✅ Encontrados {len(tarefas)} arquivos PDF para processar em {len(subpastas_matricula)} pastas.")
    print("...Iniciando processamento em paralelo...")

    # 2. Executa as tarefas em paralelo, usando todos os núcleos do processador
    with concurrent.futures.ProcessPoolExecutor() as executor:
        executor.map(processar_um_arquivo, tarefas)
    
    fim = time.time()
    duracao = fim - inicio
    print("\n✅ Processo finalizado.")
    print(f"⏱️ Tempo total de execução: {duracao:.2f} segundos.")

if __name__ == "__main__":
    main()