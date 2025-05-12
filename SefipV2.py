import os

import pandas as pd

import re

import fitz  # PyMuPDF

from tqdm import tqdm

from concurrent.futures import ProcessPoolExecutor
 
def limpar_cnpj(cnpj):

    return re.sub(r'\D', '', str(cnpj))
 
def extrair_comp(texto):

    match = re.search(r'COMP\s*[:\.\-]?\s*(\d{2})/(\d{4})', texto)

    if match:

        mes, ano = match.groups()

        return mes, ano

    return None, None
 
def processar_pdf(caminho_pdf, cnpj_original, cnpj_limpo, mes_filtro, ano_filtro):

    paginas_encontradas = []
 
    if not os.path.exists(caminho_pdf):

        return paginas_encontradas
 
    doc = fitz.open(caminho_pdf)
 
    for i in range(len(doc)):

        pagina = doc.load_page(i)

        texto = pagina.get_text()
 
        texto_limpo = limpar_cnpj(texto) if texto else ""
 
        if (cnpj_original in texto or cnpj_limpo in texto_limpo):

            mes, ano = extrair_comp(texto)

            if mes and ano and mes == mes_filtro and ano == ano_filtro:

                paginas_encontradas.append((ano, mes, caminho_pdf, i, texto.strip()))
 
    doc.close()

    return paginas_encontradas
 
def main():

    caminho_excel = r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\teste3\Filiais ativas.xlsx'
 
    pasta_pdf_base = r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023'
 
    pasta_destino =  r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\teste3'
 
    os.makedirs(pasta_destino, exist_ok=True)
 
    print("Lendo lista de CNPJs...")

    df = pd.read_excel(caminho_excel)
 
    meses_lista = [f'{i:02d}' for i in range(1, 13)]
 
    for mes_loop in meses_lista:

        mes_ano_label = f'{mes_loop}_2023'

        print(f"\n========== Iniciando processamento do mês: {mes_loop}/2023 ==========")
 
        # Listar todos PDFs do mês atual

        pasta_mes_atual = os.path.join(pasta_pdf_base, f'{mes_loop}_2023')

        caminhos_pdf_mes = [os.path.join(pasta_mes_atual, f) for f in os.listdir(pasta_mes_atual) if f.lower().endswith('.pdf')]
 
        for index, row in tqdm(df.iterrows(), total=len(df), desc=f"Processando CNPJs do mês {mes_loop}"):

            status_atual = str(row.get(mes_ano_label, '')).strip().lower()

            if status_atual == 'concluído':

                print(f"✔️  {row['CNPJ']} já concluído para {mes_loop}/2023.")

                continue
 
            cnpj_original = str(row['CNPJ'])

            cnpj_limpo = limpar_cnpj(cnpj_original)

            filial_numero = str(row['Filial'])
 
            paginas_por_comp = {}
 
            with ProcessPoolExecutor() as executor:

                futures = [

                    executor.submit(processar_pdf, caminho_pdf, cnpj_original, cnpj_limpo, mes_loop, '2023')

                    for caminho_pdf in caminhos_pdf_mes

                ]
 
                for future in futures:

                    paginas_encontradas = future.result()

                    for ano, mes, caminho_pdf, num_pagina, texto_pagina in paginas_encontradas:

                        chave = (ano, mes)

                        if chave not in paginas_por_comp:

                            paginas_por_comp[chave] = []

                        paginas_por_comp[chave].append((caminho_pdf, num_pagina, texto_pagina))
 
            if paginas_por_comp:

                nome_pasta_filial = f"Filial - {filial_numero}"

                pasta_filial = os.path.join(pasta_destino, nome_pasta_filial)

                os.makedirs(pasta_filial, exist_ok=True)
 
                for (ano, mes), paginas in paginas_por_comp.items():

                    pasta_ano = os.path.join(pasta_filial, ano)

                    os.makedirs(pasta_ano, exist_ok=True)
 
                    novo_pdf = fitz.open()

                    textos_unicos = set()
 
                    for caminho_pdf, num_pagina, texto_pagina in paginas:

                        texto_hash = hash(texto_pagina)

                        if texto_hash in textos_unicos:

                            continue

                        textos_unicos.add(texto_hash)
 
                        with fitz.open(caminho_pdf) as doc_origem:

                            novo_pdf.insert_pdf(doc_origem, from_page=num_pagina, to_page=num_pagina)
 
                    nome_pdf_saida = f"SEFIP - {mes}.pdf"

                    caminho_saida = os.path.join(pasta_ano, nome_pdf_saida)

                    novo_pdf.save(caminho_saida)

                    novo_pdf.close()
 
                    print(f"✅ {cnpj_original} salvo em {caminho_saida}")
 
                df.at[index, mes_ano_label] = "Concluído"
 
            else:

                print(f"❌ Nenhuma página encontrada para {cnpj_original} no mês {mes_loop}/2023.")

                df.at[index, mes_ano_label] = "Não encontrado"
 
            df.to_excel(caminho_excel, index=False)
 
    print("\n🚀 Processo finalizado!")

    print(f"Excel atualizado em: {caminho_excel}")
 
if __name__ == '__main__':

    main()

 