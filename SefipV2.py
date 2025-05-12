import os

import pandas as pd

import re

import fitz  # PyMuPDF

from tqdm import tqdm

from concurrent.futures import ProcessPoolExecutor
 
# Função para limpar CNPJ

def limpar_cnpj(cnpj):

    return re.sub(r'\D', '', str(cnpj))
 
# Função para extrair MÊS e ANO após "COMP:"

def extrair_comp(texto):

    match = re.search(r'COMP\s*[:\.\-]?\s*(\d{2})/(\d{4})', texto)

    if match:

        mes, ano = match.groups()

        return mes, ano

    return None, None
 
# Função para processar cada arquivo PDF e retornar páginas relevantes

def processar_pdf(caminho_pdf, cnpj_original, cnpj_limpo):

    paginas_encontradas = []

    if not os.path.exists(caminho_pdf):

        return paginas_encontradas

    doc = fitz.open(caminho_pdf)

    for i in range(len(doc)):

        pagina = doc.load_page(i)

        texto = pagina.get_text()

        texto_limpo = limpar_cnpj(texto) if texto else ""

        if cnpj_original in texto or cnpj_limpo in texto_limpo:

            mes, ano = extrair_comp(texto)

            if mes and ano:

                paginas_encontradas.append((ano, mes, caminho_pdf, i, texto.strip()))

    doc.close()

    return paginas_encontradas
 
def main():

    caminho_excel = r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\teste3\Filiais ativas 1 4 (1).xlsx'

    caminhos_pdf = [

            r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\01_2023\RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\01_2023\Relatório RE II.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\02_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\02_2023\Relatório RE 2.pdf',
       
          r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\03_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\04_2023\Relatório RE 2.pdf',
       
          r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\03_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\04_2023\Relatório RE 2.pdf',
       
          r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\05_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\05_2023\Relatório RE 2.pdf',
       
             r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\06_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\06_2023\Relatório RE 2.pdf',
       
              r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\07_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\07_2023\Relatório RE 2.pdf',
       
            r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\08_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\08_2023\Relatório RE 2.pdf',
         r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\08_2023\Relatório RE 3.pdf',
 
            r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\09_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\09_2023\Relatório RE 2.pdf',
 
 r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\10_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\10_2023\Relatório RE 2.pdf',
         r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\11_2023\Relatório RE.pdf',
 
        r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\11_2023\Relatório RE 2.pdf',
       
         r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\SeFip\2023\12_2023\Relatório RE.pdf',
    ]

       
        # Inclua mais caminhos de PDF conforme necessário

    

    pasta_destino = r'C:\Users\2160036544\Downloads\Demandas\Fiscalização\SEFIP\teste3'

    os.makedirs(pasta_destino, exist_ok=True)
 
    print("Lendo lista de CNPJs...")

    df = pd.read_excel(caminho_excel)

    log_lista = []
 
    # Processamento de cada mês

    for mes_atual in ['01_2023', '02_2023', '03_2023', '04_2023', '05_2023', '06_2023', '07_2023', '08_2023', '09_2023', '10_2023', '11_2023', '12_2023']:

        print(f"Iniciando processamento do mês: {mes_atual}")

        caminhos_pdf_mes = [pdf for pdf in caminhos_pdf if mes_atual.split('_')[0] in pdf]
 
        for index, row in tqdm(df.iterrows(), total=len(df), desc=f"Processando CNPJs do mês {mes_atual}"):

            # Verificar se já está concluído e pular

            status_atual = str(row.get(mes_atual, '')).strip().lower()

            if status_atual == 'concluído':

                continue  # pula para a próxima linha
 
            cnpj_original = str(row['CNPJ'])

            cnpj_limpo = limpar_cnpj(cnpj_original)

            filial_numero = str(row['Filial'])
 
            paginas_por_comp = {}
 
            # Usar ProcessPoolExecutor para processar os PDFs em paralelo

            with ProcessPoolExecutor() as executor:

                futures = [

                    executor.submit(processar_pdf, caminho_pdf, cnpj_original, cnpj_limpo)

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

                # Criar pasta principal para a filial, por exemplo, "Filial - 123"

                pasta_filial = os.path.join(pasta_destino, f"Filial - {filial_numero}")

                os.makedirs(pasta_filial, exist_ok=True)
 
                # Criar subpasta para o ano dentro da pasta da filial

                for (ano, mes), paginas in paginas_por_comp.items():

                    pasta_ano = os.path.join(pasta_filial, str(ano))

                    os.makedirs(pasta_ano, exist_ok=True)
 
                    novo_pdf = fitz.open()

                    textos_unicos = set()
 
                    for caminho_pdf, num_pagina, texto_pagina in paginas:

                        texto_hash = hash(texto_pagina)

                        if texto_hash in textos_unicos:

                            continue  # Página duplicada, ignora

                        textos_unicos.add(texto_hash)

                        with fitz.open(caminho_pdf) as doc_origem:

                            novo_pdf.insert_pdf(doc_origem, from_page=num_pagina, to_page=num_pagina)
 
                    # Renomear o arquivo conforme o mês encontrado no campo COMP:

                    nome_pdf_saida = f"SEFIP - {mes}.pdf"

                    caminho_saida = os.path.join(pasta_ano, nome_pdf_saida)

                    novo_pdf.save(caminho_saida)

                    novo_pdf.close()
 
                    # Log do processamento

                    log_lista.append({

                        'CNPJ': cnpj_original,

                        'Filial': filial_numero,

                        'Ano': ano,

                        'Mês': mes,

                        'Arquivo Gerado': caminho_saida,

                        'Status': 'Páginas únicas salvas'

                    })
 
                df.at[index, mes_atual] = "Concluído"

            else:

                print(f" - Nenhuma página encontrada para {cnpj_original} no mês {mes_atual}.")

                log_lista.append({

                    'CNPJ': cnpj_original,

                    'Filial': filial_numero,

                    'Ano': '',

                    'Mês': mes_atual,

                    'Arquivo Gerado': '',

                    'Status': 'Não encontrado'

                })

                df.at[index, mes_atual] = "Não encontrado"
 
            # Atualiza o arquivo Excel com o status

            df.to_excel(caminho_excel, index=False)
 
        # Salvar log do mês

        log_df = pd.DataFrame(log_lista)

        log_saida = os.path.join(pasta_destino, f'Log_Extração_{mes_atual}.xlsx')

        log_df.to_excel(log_saida, index=False)
 
        print(f"Log do mês {mes_atual} salvo em: {log_saida}")
 
    print("Processo finalizado!")

    print(f"Excel atualizado em: {caminho_excel}")
 
if __name__ == '__main__':

    main()

 