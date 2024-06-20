from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import os

def obter_dados(loja, url, xpath):
    options = webdriver.ChromeOptions()
    options.headless = True
    service = Service()

    driver = webdriver.Chrome(service=service, options=options)
    driver.get(url)
    
    try:
        elemento = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        preco = elemento.text.strip()
    except Exception as e:
        preco = "não encontrado"
    finally:
        driver.quit()
    
    return {"LOJA": loja, "PRECO": preco}

def main():
    resultado = [
        obter_dados('Amazon', 'https://www.amazon.com.br/dp/B0CYNG67P9/', '//*[@id="corePrice_feature_div"]/div/div/span[1]/span[2]/span[2]'),
        #obter_dados("Fast Shop", "https://www.fastshop.com.br/web/p/d/3005647532_PRD/samsung-smart-43-crystal-uhd-4k-43du8000-2024-painel-dynamic-crystal-color-alexa-built-in", "span.price-fraction"),
        #obter_dados("Loja oficial Samsung", "https://www.samsung.com/br/tvs/uhd-4k-tv/du8000-43-inch-crystal-uhd-4k-tizen-os-smart-tv-un43du8000gxzd/", "strong.cost-box__price-now"),
        obter_dados('Casas Bahia', 'https://www.casasbahia.com.br/samsung-smart-tv-43-quot-crystal-uhd-4k-43du8000-2024-painel-dynamic-crystal-color-alexa-built-in/p/1566004815?utm_source=zoom&utm_medium=comparadorpreco&utm_campaign=1ad9f36014c24b0699cd3f31221bf2a6', '//*[@id="product-price"]/span[1]')]

    criar_planilha(resultado)
    
def criar_planilha(resultado):
    wb = Workbook()
    ws = wb.active
    ws.append(["LOJA", "PRECO"])
    for item in resultado:
        ws.append([item["LOJA"], item["PRECO"]])
    for coluna in ws.columns:
        max_comprimento = 0
        column = coluna[0].column_letter
        for cell in coluna:
            try:
                if len(str(cell.value)) > max_comprimento:
                    max_comprimento = len(cell.value)
            except:
                pass
        largura_ajustada = (max_comprimento + 2) * 1.2
        ws.column_dimensions[column].width = largura_ajustada
    planilha = "resultados.xlsx"
    wb.save(planilha)
    os.system(f'start {planilha}')

if __name__ == "__main__":
    main()
