# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog
import os
import logging
from datetime import datetime
import cv2
import pytesseract
import fitz      # PyMuPDF
from PIL import Image # Pillow para manipular imagens
import io
import numpy as np
import shutil

# =================================================================================
# !! IMPORTANTE !! CONFIGURE SUAS INFORMAÇÕES AQUI
# =================================================================================

# 1. CAMINHO PARA O EXECUTÁVEL DO TESSERACT
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\2160036544\Downloads\tesseract-5.5.1\tesseract.exe"

# 2. MAPA DE PALAVRAS-CHAVE (SEU MAPA ORIGINAL)
palavras_chave = {
    "ASO": ["Atestado de Saúde Ocupacional", "ATESTADO DE SAÚDE OCUPACIONAL", "A.S.O."],
    "Acordo BH": ["Acordo individual de Banco de Horas", "Acordo de compensação de horas de trabalho", "ACORDO DE COMPENSAÇÃO DE HORAS DE TRABALHO", "ACORDO PARA PRORROGAÇÃO DE HORAS DE TRABALHO", "ACORDO INDIVIDUAL DE BANCO DE HORAS"],
    "Contrato de Trabalho": ["CONTRATO DE TRABALHO", "CONTRATO DE EXPERIÊNCIA", "CONTRATO DE APRENDIZAGEM"],
    "Ficha de Registro": ["FICHA DE REGISTRO", "REGISTRO DE EMPREGADO"],
    "CTPS": ["Ficha de Anotações e Atualizações da CTPS", "FICHA DE ANOTAÇÕES E ATUALIZAÇÕES DA CTPS"],
    "TRCT": ["TERMO DE RESCISÃO DO CONTRATO DE TRABALHO", "TRCT"],
    "Telegrama": ["TELEGRAMA"],
    "Comunicado de Dispensa": ["Comunicado De Dispensa", "Comunicado de Justa Causa", "Carta de Pedido de Demissão", "Comunicado de Término do Contrato de Trabalho", "Aviso de Desligamento"],
    "Aviso de Férias": ["A V I S O D E F É R I A S", "aviso de férias"],
    "FGTS": ["Extrato Analítico"],
    "RB": ["Detalhe do Pagamento - Crédito em Conta Salário", "Detalhe do Pagamento - Crédito em Conta Corrente", "Comprovante de pagamento", "COMPROVANTE DE PAGAMENTO"],
    "Termos": ["Termo", "TERMO"],
    "Aditamento": ["Aditamento do Contrato de Trabalho", "Termo Aditivo"]
}

# 3. ESTRUTURA DE ORGANIZAÇÃO DE PASTAS (SUA ESTRUTURA ORIGINAL)
organizacao = {
    "Contrato de Trabalho": ["Contrato de Trabalho", "Acordo BH", "Ficha de Registro", "Termos"],
    "CTPS": ["CTPS"],
    "TRCT Homologado": ["TRCT", "FGTS", "RB"],
    "Comunicados": ["Comunicado de Dispensa", "Telegrama"],
    "Aditamento": ["Aditamento"],
    "Atestado de Saúde Ocupacional": ["ASO"],
    "Aviso de Férias": ["Aviso de Férias"]
}


# =================================================================================
# Bloco de Configuração do Logging
# =================================================================================
# log_filename = f"analise_lote_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
# logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', datefmt="%H:%M:%S", filename=log_filename, filemode="w")
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt="%H:%M:%S")
console_handler.setFormatter(formatter)
logging.getLogger().addHandler(console_handler)

# =================================================================================
# FUNÇÕES DE ANÁLISE (COM LÓGICA PARA MÚLTIPLOS FORMATOS)
# =================================================================================

def processar_imagem_para_ocr(imagem_pil):
    """Função auxiliar que aplica filtros e OCR a uma imagem do Pillow."""
    cv_img = cv2.cvtColor(np.array(imagem_pil), cv2.COLOR_RGB2BGR)
    gray_img = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    thresh_img = cv2.threshold(gray_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return pytesseract.image_to_string(thresh_img, lang='por')

def extrair_texto_de_arquivo(caminho_arquivo):
    """
    Função principal que verifica a extensão e aplica o método de extração correto.
    """
    texto_completo = ""
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    try:
        if extensao == '.pdf':
            # Primeiro, tenta extrair texto nativo do PDF
            with fitz.open(caminho_arquivo) as doc:
                for page in doc:
                    texto_completo += page.get_text()
            
            # Se o texto for curto, aplica o OCR avançado
            if len(texto_completo.strip()) < 50:
                logging.info("   -> PDF com pouco texto, aplicando OCR avançado...")
                texto_completo = "" # Zera para preencher com OCR
                with fitz.open(caminho_arquivo) as doc:
                    for i, page in enumerate(doc):
                        pix = page.get_pixmap(dpi=300)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        texto_completo += processar_imagem_para_ocr(img) + f"\n--- Página {i+1} ---\n"
        
        elif extensao in ['.tif', '.tiff']:
            logging.info("   -> Arquivo TIF detectado, aplicando OCR...")
            with Image.open(caminho_arquivo) as img:
                for i in range(img.n_frames):
                    img.seek(i)
                    texto_completo += processar_imagem_para_ocr(img.convert("RGB")) + f"\n--- Frame {i+1} ---\n"

        elif extensao in ['.jpeg', '.jpg', '.png']:
            logging.info("   -> Arquivo de imagem detectado, aplicando OCR...")
            with Image.open(caminho_arquivo) as img:
                texto_completo += processar_imagem_para_ocr(img.convert("RGB"))

    except Exception as e:
        logging.error(f"   -> Erro ao extrair texto do arquivo {os.path.basename(caminho_arquivo)}: {e}")
        return ""
        
    return texto_completo

def identificar_categoria(texto):
    """Sua função original de verificação de categoria."""
    if not texto or not texto.strip():
        return "Nao_Identificado"
    for categoria, lista_palavras in palavras_chave.items():
        for palavra in lista_palavras:
            if palavra.lower() in texto.lower():
                logging.info(f"   -> Palavra-chave encontrada: '{palavra}'. Categoria: {categoria}")
                return categoria
    return "Nao_Identificado"


# =================================================================================
# FLUXO PRINCIPAL (ADAPTADO PARA BUSCAR MAIS TIPOS DE ARQUIVO)
# =================================================================================
def criar_pastas_e_processar(pasta_base):
    encontrados_path = os.path.join(pasta_base, "Encontrados")
    os.makedirs(encontrados_path, exist_ok=True)
    
    nao_mapeado_path = os.path.join(encontrados_path, "Categorias_Nao_Mapeadas")
    nao_identificado_path = os.path.join(encontrados_path, "Nao_Identificados")
    os.makedirs(nao_mapeado_path, exist_ok=True)
    os.makedirs(nao_identificado_path, exist_ok=True)

    # --- ALTERAÇÃO: Define as extensões de arquivo que queremos processar ---
    EXTENSOES_SUPORTADAS = ('.pdf', '.jpeg', '.jpg', '.tif', '.tiff', '.png')

    for arquivo in os.listdir(pasta_base):
        if arquivo.lower().endswith(EXTENSOES_SUPORTADAS):
            caminho_arquivo = os.path.join(pasta_base, arquivo)
            logging.info(f"🔍 Processando: {arquivo}")

            # Chama a nova função de extração que lida com todos os formatos
            texto = extrair_texto_de_arquivo(caminho_arquivo)
            categoria = identificar_categoria(texto)
            logging.info(f"   -> Categoria final: {categoria}")

            destino = None
            if categoria == "Nao_Identificado":
                destino = nao_identificado_path
            else:
                for pasta_principal, subpastas in organizacao.items():
                    if categoria in subpastas:
                        destino = os.path.join(encontrados_path, pasta_principal, categoria)
                        os.makedirs(destino, exist_ok=True)
                        break
                if destino is None:
                    logging.warning(f"⚠️ Categoria '{categoria}' encontrada, mas não mapeada. Movendo para pasta de não mapeados.")
                    destino = nao_mapeado_path
            
            nome_final = arquivo
            contador = 1
            while os.path.exists(os.path.join(destino, nome_final)):
                nome_base, extensao = os.path.splitext(arquivo)
                nome_final = f"{nome_base} ({contador}){extensao}"
                contador += 1
            
            logging.info(f"✅ Movendo para: {os.path.relpath(destino, pasta_base)}")
            shutil.move(caminho_arquivo, os.path.join(destino, nome_final))

    logging.info("\n🚀 Processamento concluído!")

if __name__ == "__main__":
    try:
        # Inicia a interface gráfica para pedir a pasta
        root = tk.Tk()
        root.withdraw() 
        logging.info("Por favor, selecione a pasta do lote que contém os arquivos (PDF, JPG, TIF).")
        pasta_lote = filedialog.askdirectory(title="Selecione a pasta do lote")

        if pasta_lote:
            logging.info(f"Pasta selecionada: {pasta_lote}")
            criar_pastas_e_processar(pasta_lote)
        else:
            logging.warning("Nenhuma pasta selecionada. Encerrando o programa.")
            
    except Exception as e:
        logging.error(f"Ocorreu um erro fatal: {e}", exc_info=True)
    finally:
        input("\nPressione Enter para fechar...")