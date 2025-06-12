import os
import json
import re  # Módulo para o filtro de limpeza
import fitz  # PyMuPDF
import yake
import pytesseract
from PIL import Image
import io
from tqdm import tqdm # Para a barra de progresso

# --- CONFIGURAÇÃO DO TESSERACT COM O SEU CAMINHO ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\2160036544\Downloads\tesseract-5.5.1\tesseract.exe'

# ==============================================================================
# FUNÇÃO DE FILTRO PARA LIMPEZA DOS RESULTADOS
# ==============================================================================
def filtrar_palavras_chave_ruidosas(lista_palavras: list) -> list:
    """Filtra uma lista de palavras-chave para remover resultados de OCR de baixa qualidade."""
    palavras_filtradas = []
    for palavra in lista_palavras:
        if len(palavra.replace(" ", "")) < 3:
            continue
        if not re.search(r'[aeiouáéíóúâêôãõ]', palavra, re.IGNORECASE):
            continue
        if re.search(r'(.)\1{3,}', palavra):
            continue
        palavras_filtradas.append(palavra)
    return palavras_filtradas

# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO OTIMIZADAS
# ==============================================================================

def extrair_texto_do_pdf(caminho_pdf: str, pbar_sub) -> str:
    """Extrai texto de um PDF, usando OCR com DPI aumentado para melhor qualidade."""
    texto_completo = ""
    try:
        doc = fitz.open(caminho_pdf)
        for i, pagina in enumerate(doc):
            pbar_sub.set_description(f"Lendo pág {i+1}/{len(doc)}")
            texto_pagina = pagina.get_text("text")
            if len(texto_pagina.strip()) < 50:
                try:
                    # AUMENTO DO DPI PARA MELHORAR A QUALIDADE DA IMAGEM PARA O OCR
                    pix = pagina.get_pixmap(dpi=2000)
                    img_bytes = pix.tobytes("png")
                    imagem = Image.open(io.BytesIO(img_bytes))
                    texto_ocr = pytesseract.image_to_string(imagem, lang='por')
                    texto_completo += texto_ocr + "\n"
                except Exception as ocr_error:
                    # print(f"  [AVISO] Erro de OCR na página {i+1}: {ocr_error}")
                    pass
            else:
                texto_completo += texto_pagina + "\n"
        doc.close()
    except Exception as e:
        print(f"  [ERRO] Não foi possível ler o arquivo '{caminho_pdf}': {e}")
        return None
    return texto_completo

def extrair_palavras_chave(texto: str, max_palavras: int = 20) -> list:
    """Extrai as palavras-chave de um texto usando YAKE!."""
    if not texto or len(texto.strip()) == 0:
        return []
    extrator = yake.KeywordExtractor(lan="pt", n=3, dedupLim=0.9, top=max_palavras, features=None)
    palavras_chave_com_score = extrator.extract_keywords(texto)
    return [chave[0] for chave in palavras_chave_com_score]

# ==============================================================================
# FUNÇÃO PRINCIPAL DE PROCESSAMENTO COM FILTRO
# ==============================================================================

def processar_pasta_raiz(pasta_raiz: str) -> dict:
    """Varre uma pasta raiz, extrai e FILTRA palavras-chave de cada PDF."""
    todos_os_arquivos_pdf = []
    for root, _, files in os.walk(pasta_raiz):
        for file in files:
            if file.lower().endswith('.pdf'):
                todos_os_arquivos_pdf.append(os.path.join(root, file))

    if not todos_os_arquivos_pdf:
        print("Nenhum arquivo PDF encontrado na pasta especificada.")
        return {}
        
    print(f"Encontrados {len(todos_os_arquivos_pdf)} arquivos PDF. Iniciando extração...")
    
    resultados_finais = {}
    
    with tqdm(total=len(todos_os_arquivos_pdf), desc="Progresso Geral") as pbar_main:
        for caminho_pdf in todos_os_arquivos_pdf:
            pbar_main.set_description(f"Processando: {os.path.basename(caminho_pdf)}")
            
            with tqdm(total=1, leave=False) as pbar_sub:
                texto = extrair_texto_do_pdf(caminho_pdf, pbar_sub)
                if texto:
                    palavras_brutas = extrair_palavras_chave(texto)
                    # APLICANDO O FILTRO DE LIMPEZA
                    palavras_limpas = filtrar_palavras_chave_ruidosas(palavras_brutas)
                    
                    if palavras_limpas:
                        resultados_finais[caminho_pdf] = palavras_limpas
            
            pbar_main.update(1)

    return resultados_finais

# ==============================================================================
# FUNÇÃO PARA GERAR AS SAÍDAS (Relatório e JSON)
# ==============================================================================

def gerar_saidas(resultados: dict, pasta_saida: str):
    """Gera o arquivo JSON e o relatório de texto com base nos resultados."""
    if not resultados:
        print("\nNenhum resultado válido encontrado para gerar as saídas.")
        return

    os.makedirs(pasta_saida, exist_ok=True)

    caminho_json = os.path.join(pasta_saida, "banco_de_dados.json")
    with open(caminho_json, 'w', encoding='utf-8') as f:
        json.dump(resultados, f, ensure_ascii=False, indent=4)
    print(f"\n✅ Banco de dados JSON salvo em: {caminho_json}")

    caminho_relatorio = os.path.join(pasta_saida, "relatorio.txt")
    relatorio_agrupado = {}
    for caminho_arquivo, palavras in resultados.items():
        pasta = os.path.dirname(caminho_arquivo)
        nome_arquivo = os.path.basename(caminho_arquivo)
        if pasta not in relatorio_agrupado:
            relatorio_agrupado[pasta] = []
        relatorio_agrupado[pasta].append((nome_arquivo, palavras))

    with open(caminho_relatorio, 'w', encoding='utf-8') as f:
        f.write("RELATÓRIO DE PALAVRAS-CHAVE POR PASTA\n" + "=" * 40 + "\n\n")
        for pasta, arquivos in sorted(relatorio_agrupado.items()):
            f.write(f"PASTA: {pasta}\n" + "-" * len(f"PASTA: {pasta}") + "\n")
            for nome_arquivo, palavras in arquivos:
                f.write(f"  -> ARQUIVO: {nome_arquivo}\n")
                if palavras:
                    f.write(f"     PALAVRAS-CHAVE: {', '.join(palavras)}\n\n")
                else: # Caso um arquivo não tenha palavras-chave válidas após o filtro
                    f.write(f"     PALAVRAS-CHAVE: (Nenhuma palavra-chave válida encontrada)\n\n")
            f.write("\n")
    print(f"✅ Relatório de texto salvo em: {caminho_relatorio}")


# ==============================================================================
# EXECUÇÃO PRINCIPAL COM OS SEUS CAMINHOS
# ==============================================================================

if __name__ == "__main__":
    # --- CAMINHOS CONFIGURADOS CONFORME SOLICITADO ---
    pasta_raiz_documentos = r"C:\Users\2160036544\Downloads\EXEMPLOS DE DOCUMENTOS"
    pasta_de_saida = r"C:\Users\2160036544\Downloads\test"
    
    if not os.path.isdir(pasta_raiz_documentos):
        print(f"[ERRO FATAL] A pasta de documentos '{pasta_raiz_documentos}' não existe.")
    else:
        resultados = processar_pasta_raiz(pasta_raiz_documentos)
        gerar_saidas(resultados, pasta_de_saida)
        print("\nProcesso concluído!")