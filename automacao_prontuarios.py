import os
import shutil
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract


# === CAMINHO PARA A PASTA DO LOTE ===
PASTA_LOTE = r"C:\Users\2160036544\Downloads\SavedDocument\lote1\1221078-IRACI GUEDES DE OLIVEIRA BITARAES SILVA"

# === CAMINHO DO TESSERACT ===
# Altere para o caminho correto no seu computador
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\2160036544\Downloads\tesseract-5.5.1\tesseract.exe"

# === CAMINHO DO POPPLER (MODIFICAÇÃO 1) ===
# ### MODIFICAÇÃO ###
# Altere para o caminho da pasta 'bin' do Poppler no seu computador
POPPLER_PATH = r"C:\Users\2160036544\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"


# === MAPA DE PALAVRAS-CHAVE ===
palavras_chave = {
    "ASO": [
        "Atestado de Saúde Ocupacional", "Atestado de saúde ocupacional", "atestado de saúde ocupacional",
        "ATESTADO DE SAÚDE OCUPACIONAL", "A.S.O."
    ],
    "Acordo BH": [
        "Acordo individual de Banco de Horas", "acordo individual de banco de horas",
        "Acordo de compensação de horas de trabalho", "ACORDO DE COMPENSAÇÃO DE HORAS DE TRABALHO",
        "ACORDO PARA PRORROGAÇÃO DE HORAS DE TRABALHO", "Acordo Para Prorrogação de Horas de trabalho",
        "ACORDO INDIVIDUAL DE BANCO DE HORAS"
    ],
    "Contrato de Trabalho": [
        "CONTRATO DE TRABALHO", "Contrato de Trabalho", "CONTRATO DE EXPERIÊNCIA", "Contrato de Experiência",
        "CONTRATO DE APRENDIZAGEM", "Contrato de Aprendizagem"
    ],
    "Ficha de Registro": [
        "FICHA DE REGISTRO", "Ficha de Registro", "REGISTRO DE EMPREGADO", "Registro de Empregado"
    ],
    "CTPS": [
        "Ficha de Anotações e Atualizações da CTPS", "FICHA DE ANOTAÇÕES E ATUALIZAÇÕES DA CTPS"
    ],
    "TRCT": [
        "TERMO DE RESCISÃO DO CONTRATO DE TRABALHO", "Termo de Rescisão do Contrato de Trabalho", "TRCT"
    ],
    "Telegrama": [
        "TELEGRAMA", "Telegrama"
    ],
    "Comunicado de Dispensa": [
        "Comunicado De Dispensa", "Comunicado de Justa Causa", "Carta de Pedido de Demissão",
        "Comunicado de Término do Contrato de Trabalho", "Aviso de Desligamento", "COMUNICADO DE DISPENSA", "COMUNICACAO DE JUSTA CAUSA", "CARTA DE PEDIDO DE DEMISSÃO",
        "COMUNICADO DE TÉRMINO DO CONTRATO DE TRABALHO", "AVISO DE DESLIGAMENTO"
    ],
    "Aviso de Férias": [
        "A V I S O D E F É R I A S", "a v i s o d e f é r i a s", "aviso de férias", "Aviso de férias"
    ],
    "FGTS": [
        "Extrato Analítico", "EXTRATO ANALÍTICO", "extrato analítico"
    ],
    "RB": [
        "Detalhe do Pagamento - Crédito em Conta Salário", "Detalhe do Pagamento - Crédito em Conta Corrente",
        "Comprovante de pagamento", "Comprovante do Pagamento", "COMPROVANTE DE PAGAMENTO",
        "COMPROVANTE DO PAGAMENTO"
    ],
    "Termos": [
        "Termo", "TERMO"
    ],
    "Aditamento": [
        "Aditamento do Contrato de Trabalho", "Termo Aditivo", "TERMO ADITIVO", "ADITAMENTO DO CONTRATO DE TRABALHO"
    ]
}


# === ORGANIZAÇÃO DAS PASTAS ===
organizacao = {
    "Contrato de Trabalho": ["Acordo BH", "Ficha de Registro", "Termos"],
    "CTPS": ["CTPS"],
    "TRCT Homologado": ["TRCT", "FGTS", "RB"],
    "Comunicados": ["Comunicado de Dispensa", "Telegrama"],
    "Aditamento": ["Aditamento"],
    "Atestado de Saúde Ocupacional": ["ASO"],
    "Aviso de Férias": ["Aviso de Férias"]
}


# === FUNÇÃO PARA EXTRAIR TEXTO DO PDF ===
def extrair_texto(pdf_path):
    texto = ""

    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            texto += page.get_text()
        doc.close()
    except:
        texto = ""

    # Se não encontrar texto, aplica OCR
    if texto.strip() == "":
        try:
            # ### MODIFICAÇÃO 2 ###
            # Adicionado o parâmetro 'poppler_path' para indicar onde está o Poppler
            imagens = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
            for img in imagens:
                texto += pytesseract.image_to_string(img, lang='por')
        except Exception as e:
            print(f"Erro no OCR do arquivo {pdf_path}: {e}")

    return texto


# === FUNÇÃO PARA VERIFICAR CATEGORIA ===
def identificar_categoria(texto):
    for categoria, lista_palavras in palavras_chave.items():
        for palavra in lista_palavras:
            if palavra.lower() in texto.lower():
                return categoria
    return None


# === CRIA ESTRUTURA DE PASTAS ===
def criar_pastas(base):
    encontrados = os.path.join(base, "Encontrados")
    os.makedirs(encontrados, exist_ok=True)

    for pasta_principal, subpastas in organizacao.items():
        pasta_principal_path = os.path.join(encontrados, pasta_principal)
        os.makedirs(pasta_principal_path, exist_ok=True)

        for sub in subpastas:
            sub_path = os.path.join(pasta_principal_path, sub)
            os.makedirs(sub_path, exist_ok=True)


# === INÍCIO DO PROCESSO ===
def executar_organizacao(pasta_base):
    criar_pastas(pasta_base)

    for arquivo in os.listdir(pasta_base):
        if arquivo.lower().endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_base, arquivo)
            print(f"🔍 Processando: {arquivo}")

            texto = extrair_texto(caminho_pdf)
            categoria = identificar_categoria(texto)

            if categoria:
                destino = None
                for pasta_principal, subpastas in organizacao.items():
                    if categoria in subpastas:
                        destino = os.path.join(pasta_base, "Encontrados", pasta_principal, categoria)
                        break
                if destino is None:
                    print(f"⚠️ Categoria {categoria} não encontrada na organização.")
                    continue

                nome_final = arquivo
                contador = 1
                while os.path.exists(os.path.join(destino, nome_final)):
                    nome_base = os.path.splitext(arquivo)[0]
                    extensao = os.path.splitext(arquivo)[1]
                    nome_final = f"{nome_base} - {contador}{extensao}"
                    contador += 1

                shutil.move(caminho_pdf, os.path.join(destino, nome_final))
                print(f"✅ Movido para: {destino}")
            else:
                print(f"❌ Nenhuma palavra-chave encontrada em: {arquivo} — Mantido na pasta raiz.")

    print("\n🚀 Processamento concluído!")


# === EXECUTAR ===
if __name__ == "__main__":
    executar_organizacao(PASTA_LOTE)