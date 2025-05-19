import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 1) Carrega a planilha
df = pd.read_excel(r"C:\Users\2160036544\Downloads\Demandas\Content\Base_ativos_16.05.xlsx")

# 2) Configura o Edge WebDriver (é preciso ter o msedgedriver compatível no PATH)
options = webdriver.EdgeOptions()
options.use_chromium = True
# options.add_argument("--headless")  # se quiser rodar sem GUI
driver = webdriver.Edge(service=EdgeService(), options=options)

# 3) URL de destino
URL = "https://documentos/bahia/gateway"

for idx, row in df.iterrows():
    # supondo que "CurrentItem[6]" era a coluna de status; aqui verificamos pelo nome ou pelo índice 5
    if pd.isna(row.iloc[5]):
        empresa   = str(row.iloc[0])
        matricula = str(row.iloc[2])

        # Navega
        driver.get(URL)

        wait = WebDriverWait(driver, 10)
        # Aguarda o campo Empresa
        emp_field = wait.until(EC.element_to_be_clickable((By.ID, "U01_CD_EMPGCB")))
        emp_field.clear()
        emp_field.send_keys(empresa)

        # Aguarda o campo Matrícula
        mat_field = wait.until(EC.element_to_be_clickable((By.ID, "U01_CD_FUN")))
        mat_field.clear()
        mat_field.send_keys(matricula)

        # Aguarda e clica no botão de processar
        btn = wait.until(EC.element_to_be_clickable((By.ID, "NM_BOT_PRC")))
        btn.click()

        # se houver alguma confirmação ou resultado, você pode aguardar aqui:
        # wait.until(EC.text_to_be_present_in_element((By.ID, "algum_id"), "Esperado"))

        # opcional: escreva na planilha que foi processado
        df.at[idx, df.columns[5]] = "Processado"

        # breve pausa entre iterações
        time.sleep(1)

# 4) Salva a planilha com o status atualizado
df.to_excel(r"C:\Users\2160036544\Downloads\Demandas\Content\Base_ativos_16.05_processed.xlsx", index=False)

driver.quit()
