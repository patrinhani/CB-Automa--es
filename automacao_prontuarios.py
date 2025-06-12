import os
import shutil
import re
import fitz
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_path

Image.MAX_IMAGE_PIXELS = None

# ----------------------- CONFIGURAÇÕES -----------------------

CAMINHO_LOTE = r"C:\Users\2160036544\Downloads\SavedDocument\lote1\1221078-IRACI GUEDES DE OLIVEIRA BITARAES SILVA"
PASTA_DEBUG = os.path.join(CAMINHO_LOTE, "DEBUG_IMAGENS")
os.makedirs(PASTA_DEBUG, exist_ok=True)

PALAVRAS_CHAVE = {
    "ASO A": ["exame admissional"], "ASO D": ["exame demissional"], "ASO P": ["exame periódico"],
    "Aditamentos do Contrato de trabalho": ["aditamento do contrato de trabalho", "termo aditivo"],
    "Contrato de Trabalho": ["contrato de trabalho", "contrato de experiencia", "acordo de compensação de horas", "acordo individual de banco de horas", "termo de ciência e aceite", "termo de responsabilidade", "registro de empregado", "termo de compromisso"],
    "Formulário de atualização de CTPS": ["ficha de anotações e atualizações da ctps"],
    "TRCT Homologado": ["extrato fgts", "guia do seguro", "extrato analítico", "guia de recolhimento rescisório", "credito em conta", "requerimento do seguro desemprego", "trct"],
    "Comunicado de Aviso de Férias": ["aviso de férias"],
    "Comunicado De Dispensa": ["aviso prévio", "justa causa", "pedido de demissão", "término do contrato", "telegramas", "comprovante de envio de torpedo", "comunicado de despensa"]
}
TIPOS_ASO = {
    "ASO A": ["admissional"], "ASO D": ["demissional"], "ASO P": ["periódico"],
    "Retorno ao Trabalho": ["retorno ao trabalho"], "Mudança de Risco Ocupacional": ["mudança de risco", "mudança de função"]
}

# ---------------------- FUNÇÕES AUXILIARES ----------------------
# As funções corrigir_inclinacao e extrair_texto permanecem as mesmas
def corrigir_inclinacao(imagem_cv):
    cinza = cv2.cvtColor(imagem_cv, cv2.COLOR_BGR2GRAY if len(imagem_cv.shape) == 3 else cv2.COLOR_RGB2GRAY)
    cinza = cv2.bitwise_not(cinza)
    coords = np.column_stack(np.where(cinza > 0))
    if len(coords) < 10: return imagem_cv
    rect = cv2.minAreaRect(coords)
    angulo = rect[-1]
    if angulo < -45: angulo = -(90 + angulo)
    else: angulo = -angulo
    (h, w) = imagem_cv.shape[:2]
    centro = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(centro, angulo, 1.0)
    rotacionada = cv2.warpAffine(imagem_cv, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return rotacionada

def extrair_texto(pdf_path):
    try:
        texto_extraido = ""
        caminho_poppler = r'C:\Users\2160036544\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin'
        with fitz.open(pdf_path) as doc:
            for page in doc:
                texto_extraido += page.get_text("text", sort=True).lower()
        
        if len(texto_extraido.strip()) < 50:
            print(f"⚠️ Texto curto detectado. Tentando OCR para {os.path.basename(pdf_path)}...")
            texto_ocr = ""
            imagens = convert_from_path(pdf_path, dpi=300, poppler_path=caminho_poppler)
            config_tesseract = '--oem 3 --psm 6'
            
            for i, pil_image in enumerate(imagens):
                nome_base_arquivo = os.path.splitext(os.path.basename(pdf_path))[0]
                imagem_cv = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
                imagem_alinhada = corrigir_inclinacao(imagem_cv)
                imagem_cinza = cv2.cvtColor(imagem_alinhada, cv2.COLOR_BGR2GRAY)
                imagem_processada = cv2.adaptiveThreshold(imagem_cinza, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                
                caminho_debug_img = os.path.join(PASTA_DEBUG, f"DEBUG_{nome_base_arquivo}_pagina_{i+1}.png")
                cv2.imwrite(caminho_debug_img, imagem_processada)

                texto_ocr += pytesseract.image_to_string(imagem_processada, lang='por', config=config_tesseract).lower()
            texto_extraido += "\n" + texto_ocr
        return texto_extraido
    except Exception as e:
        print(f"❌ Erro ao extrair texto de {os.path.basename(pdf_path)}: {e}")
        return ""

def identificar_tipo_aso(texto):
    for tipo, palavras in TIPOS_ASO.items():
        for palavra in palavras:
            padrao = fr"\([xX\-]\)\s*{palavra}"
            if re.search(padrao, texto):
                return tipo
    return None

# FUNÇÃO DE CLASSIFICAÇÃO COM A NOVA REGRA
def classificar_documento(matricula, pasta_lote_anexos, pdf_path, texto):
    encontrados = []
    
    # 1. Lógica de alta precisão para ASO continua a mesma.
    if "atestado de saúde ocupacional" in texto:
        tipo_aso_especifico = identificar_tipo_aso(texto)
        if tipo_aso_especifico:
            encontrados.append(tipo_aso_especifico)

    # 2. Se nenhum ASO foi encontrado, aplica a busca geral com a NOVA REGRA.
    if not encontrados:
        for subpasta, palavras in PALAVRAS_CHAVE.items():
            # A lógica de ASO já foi tratada, pulamos aqui para não reclassificar.
            if subpasta.startswith("ASO"):
                continue

            # --- NOVA LÓGICA DE MÚLTIPLAS PALAVRAS-CHAVE ---
            
            # Conta quantas palavras-chave desta categoria foram encontradas no texto.
            num_palavras_encontradas = sum(1 for p in palavras if p in texto)
            
            # Pega o número de palavras-chave definidas para esta categoria.
            num_palavras_definidas = len(palavras)

            # Aplica a regra:
            # - Se a categoria só tem 1 palavra-chave, 1 acerto é suficiente.
            # - Se a categoria tem mais de 1, precisamos de pelo menos 2 acertos.
            if num_palavras_definidas == 1 and num_palavras_encontradas >= 1:
                encontrados.append(subpasta)
            elif num_palavras_definidas > 1 and num_palavras_encontradas >= 2:
                encontrados.append(subpasta)
            # --- FIM DA NOVA LÓGICA ---
    
    # O resto da função para criar os arquivos e pastas permanece o mesmo.
    if encontrados:
        # A lógica de múltiplos acertos pode, em teoria, classificar um doc em mais de uma categoria.
        # A linha abaixo garante que cada documento seja copiado apenas uma vez por categoria encontrada.
        for tipo in set(encontrados): 
            pasta_destino_tipo = os.path.join(pasta_lote_anexos, tipo)
            os.makedirs(pasta_destino_tipo, exist_ok=True)
            base_nome_arquivo = f"{matricula} - {tipo}"
            contador = 1
            novo_nome_arquivo = f"{base_nome_arquivo} - {contador}.pdf"
            caminho_destino_final = os.path.join(pasta_destino_tipo, novo_nome_arquivo)
            while os.path.exists(caminho_destino_final):
                contador += 1
                novo_nome_arquivo = f"{base_nome_arquivo} - {contador}.pdf"
                caminho_destino_final = os.path.join(pasta_destino_tipo, novo_nome_arquivo)
            shutil.copy2(pdf_path, caminho_destino_final)
            print(f"📎 Classificado: {os.path.basename(pdf_path)} ➤ {caminho_destino_final}")
    else:
        # O log de debug para arquivos não classificados continua ativo.
        print(f"🟡 Ignorado (sem palavras-chave suficientes): {os.path.basename(pdf_path)}")
        print("--------------------------- DEBUG INFO ---------------------------")
        if not texto.strip(): print(">>> MOTIVO: Nenhum texto foi extraído do arquivo.")
        else: print(f">>> TEXTO EXTRAÍDO (amostra):\n---\n{texto[:600].strip()}\n---")
        print("------------------------- FIM DEBUG INFO -------------------------")

# As funções main e if __name__ == '__main__' permanecem as mesmas
def main():
    print("🚀 Iniciando organização e classificação de lote...")
    if not os.path.exists(CAMINHO_LOTE):
        print(f"❌ ERRO: O caminho do lote não foi encontrado: {CAMINHO_LOTE}")
        return
    pastas_matriculas = [d for d in os.listdir(CAMINHO_LOTE) if os.path.isdir(os.path.join(CAMINHO_LOTE, d)) and not d.endswith(("Lote Anexos", "DEBUG_IMAGENS"))]
    if not pastas_matriculas:
        print("⚠️ Nenhuma pasta de matrícula encontrada no caminho especificado.")
        return
    for pasta_mat in pastas_matriculas:
        try:
            matricula = pasta_mat.split('-')[0].strip()
            caminho_matricula_origem = os.path.join(CAMINHO_LOTE, pasta_mat)
            print(f"\n📁 Processando Matrícula: {matricula} (Pasta: {pasta_mat})")
            pasta_lote_anexos = os.path.join(CAMINHO_LOTE, f"{matricula} - Lote Anexos")
            os.makedirs(pasta_lote_anexos, exist_ok=True)
            arquivos = [arq for arq in os.listdir(caminho_matricula_origem) if arq.lower().endswith(".pdf")]
            if not arquivos:
                print("   ⚠️ Nenhum arquivo PDF encontrado nesta pasta.")
                continue
            for arquivo in arquivos:
                caminho_pdf = os.path.join(caminho_matricula_origem, arquivo)
                texto = extrair_texto(caminho_pdf)
                if texto:
                    classificar_documento(matricula, pasta_lote_anexos, caminho_pdf, texto)
        except Exception as e:
            print(f"❌ Ocorreu um erro inesperado ao processar a pasta {pasta_mat}: {e}")
    print("\n✅ Conclusão: Análise e organização finalizadas.")

if __name__ == "__main__":
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\2160036544\Downloads\tesseract-5.5.1\tesseract.exe'
    main()