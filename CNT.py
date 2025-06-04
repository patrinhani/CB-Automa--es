import os
import shutil
import fitz # Importado para extrair texto de PDFs
import pandas as pd
from pdf2image import convert_from_path
import easyocr
import numpy as np
from fuzzywuzzy import fuzz
import unicodedata
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
 
# --- Configuração de Logging ---
# Configura o sistema de log para mostrar informações no console.
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
 
# --- Configurações Principais ---
# Inicializa o leitor EasyOCR uma única vez para português, sem usar GPU.
# Isso economiza tempo e recursos.
reader = easyocr.Reader(['pt'], gpu=False)
 
# Caminho para o executável Poppler. Essencial para o pdf2image funcionar.
poppler_path = r"C:\Users\2160036544\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"
 
# Lista de palavras-chave a serem buscadas nos documentos.
palavras_chave = [
    "adesão de beneficios no processo admissional",
    "solicitação de transporte",
    "inclusão, solicitação e cancelamento vale transporte via varejo",
    "Extrato de beneficios",
    "Vale transporte",
    "Adesao",
    "RH - CONTRATO DE TRABALHO - ADESAO",
    "EXCLUSAO DO VALE TRANSPORTE",
    "RH - ANEXOS",
    "RH - BENEFICIOS",
    "RH - CONTRATO DE TRABALHO - ADESAO",
    "Adesão de Benefícios no Processo Admissional Via Varejo e VVLOG",
    "Alteração, Inclusão e Exclusão de vale transporte",
]
 
# Caminhos das pastas e arquivo de saída.
pasta_principal = r"C:\Users\2160036544\Downloads\Arquivos_Descompactados\2160036544---N-64342692----30-05-2025-16-56-01"
saida_excel = r"C:\Users\2160036544\Downloads\Arquivos_Descompactados"
 
# Cria a pasta para PDFs encontrados se ela não existir.
pasta_encontrados = os.path.join(pasta_principal, "Encontrados")
os.makedirs(pasta_encontrados, exist_ok=True)
 
# Lista para armazenar os resultados do processamento de cada PDF.
resultados = []
 
# --- Funções Auxiliares ---
def normalizar(texto):
    """
    Normaliza o texto removendo acentos, convertendo para minúsculas
    e retirando caracteres especiais.
    """
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    texto = re.sub(r'[^a-zA-Z0-9\s]', '', texto)
    return texto.lower().strip()
 
def contem_palavra_chave(texto_normalizado):
    """
    Verifica se o texto normalizado contém alguma das palavras-chave
    usando a correspondência fuzzy (similaridade parcial).
    Um limiar de 85 significa que 85% do texto da chave deve corresponder.
    """
    for chave in palavras_chave:
        chave_norm = normalizar(chave)
        if fuzz.partial_ratio(chave_norm, texto_normalizado) >= 85:
            return True
    return False
 
def extrair_texto_fitz(caminho_pdf):
    """
    Tenta extrair texto diretamente de um PDF usando PyMuPDF (Fitz).
    É muito mais rápido para PDFs com camada de texto.
    """
    texto = ""
    try:
        doc = fitz.open(caminho_pdf)
        for page in doc:
            texto += page.get_text() # Extrai texto da página
        doc.close()
        return normalizar(texto)
    except Exception as e:
        # Se ocorrer um erro (ex: PDF corrompido, protegido), registre e retorne None.
        # Nível DEBUG para não poluir o console com erros esperados de PDFs sem texto.
        logging.debug(f"Erro ao extrair texto com Fitz de {caminho_pdf}: {e}")
        return None # Indica que Fitz não conseguiu extrair texto ou houve um erro
 
def extrair_texto_easyocr(caminho_pdf):
    """
    Extrai texto de um PDF usando EasyOCR após converter as páginas em imagens.
    Usado para PDFs que são primariamente imagens (escaneados).
    O DPI (Dots Per Inch) afeta a qualidade da imagem e o consumo de recursos.
    Valores como 150 ou 120 podem ser mais leves que 200, mas teste a precisão.
    """
    texto_total = ""
    try:
        # Tente 150 ou 120 DPI para menos consumo de recursos,
        # mas verifique se a precisão do OCR se mantém.
        imagens = convert_from_path(caminho_pdf, dpi=150, poppler_path=poppler_path)
        for img in imagens:
            # Converte a imagem para um array numpy e passa para o EasyOCR.
            resultado = reader.readtext(np.array(img), detail=0)
            texto_total += " ".join(resultado) + " "
    except Exception as e:
        logging.error(f"Erro na conversão OCR de {caminho_pdf}: {e}")
    return normalizar(texto_total)
 
def processar_pdf(caminho_pdf):
    """
    Função principal para processar um único PDF.
    Primeiro tenta extração de texto com Fitz, se falhar, usa EasyOCR.
    """
    try:
        nome_arquivo = os.path.basename(caminho_pdf)
        pasta_matricula = os.path.basename(os.path.dirname(caminho_pdf))
 
        # --- Otimização: Filtro Inicial por Nome do Arquivo ---
        # Ignora PDFs que provavelmente não contêm as palavras-chave com base no nome.
        # Isso evita processamento pesado (OCR) em arquivos irrelevantes.
        if not any(x in normalizar(nome_arquivo) for x in ["contrato de trabalho - adesao", "extrato de beneficios", "vale transporte", "beneficio", "adesao", "rh - anexos", "rh - beneficios","rh - contrato de trabalho - adesao"]):
            return [pasta_matricula, "Ignorado (Filtrado por Nome)", caminho_pdf]
 
        # --- Estratégia Híbrida: Tentar Fitz Primeiro ---
        texto_extraido_fitz = extrair_texto_fitz(caminho_pdf)
        if texto_extraido_fitz and contem_palavra_chave(texto_extraido_fitz):
            status = "Encontrado (via Fitz)"
            destino = os.path.join(pasta_encontrados, nome_arquivo)
            shutil.copy2(caminho_pdf, destino)
            logging.info(f"Copiado: {caminho_pdf} para {destino}")
            return [pasta_matricula, status, caminho_pdf]
 
        # --- Recorrer ao EasyOCR se Fitz falhar ou não encontrar ---
        # Isso significa que o PDF é uma imagem ou o texto não foi detectável por Fitz.
        logging.info(f"Fitz não encontrou ou não conseguiu ler, usando OCR para: {nome_arquivo}")
        texto_extraido_ocr = extrair_texto_easyocr(caminho_pdf)
        achou_ocr = contem_palavra_chave(texto_extraido_ocr)
        status = "Encontrado (via OCR)" if achou_ocr else "Não encontrado"
 
        if achou_ocr:
            destino = os.path.join(pasta_encontrados, nome_arquivo)
            shutil.copy2(caminho_pdf, destino)
            logging.info(f"Copiado: {caminho_pdf} para {destino}")
 
        return [pasta_matricula, status, caminho_pdf]
 
    except Exception as e:
        logging.error(f"Erro geral ao processar {caminho_pdf}: {e}", exc_info=True) # exc_info=True para detalhes do erro
        return [os.path.basename(os.path.dirname(caminho_pdf)), f"Erro: {e}", caminho_pdf]
 
# --- Função Principal de Busca ---
def buscar_com_threads(pasta):
    """
    Percorre o diretório, encontra todos os PDFs e os processa
    usando um pool de threads para paralelização.
    """
    pdfs = []
    for raiz, _, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            if arquivo.lower().endswith('.pdf'):
                pdfs.append(os.path.join(raiz, arquivo))
 
    logging.info(f"🔍 Total de PDFs a processar: {len(pdfs)}")
 
    # Otimização: Define o número de threads para um valor ideal para o seu CPU (4 núcleos).
    # Isso evita sobrecarga e mantém o sistema responsivo.
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(processar_pdf, pdf) for pdf in pdfs]
        for i, future in enumerate(as_completed(futures)):
            try:
                resultado = future.result()
                resultados.append(resultado)
                # Relata o progresso a cada 10 arquivos ou ao final.
                if (i + 1) % 10 == 0 or (i + 1) == len(pdfs):
                    logging.info(f"Progresso: {i + 1}/{len(pdfs)} PDFs processados. Último: {resultado[0]} - {resultado[1]}")
            except Exception as e:
                logging.error(f"Erro ao obter resultado de uma thread: {e}", exc_info=True)
 
# --- Execução do Script ---
if __name__ == "__main__":
    logging.info("Iniciando a busca otimizada de documentos...")
    buscar_com_threads(pasta_principal)
 
    # Cria um DataFrame do pandas com os resultados e salva em Excel.
    df = pd.DataFrame(resultados, columns=["Matrícula", "Status", "Caminho"])
    df.to_excel(saida_excel, index=False)
 
    logging.info("✅ Busca otimizada finalizada.")
    logging.info(f"📁 Planilha de resultados gerada em: {saida_excel}")