from selenium import webdriver #criar o navegador
from selenium.webdriver.common.keys import Keys#permite escrever no navegador
from selenium.webdriver.common.by import By#permite você selecionar itens no navegador
#abrir o navegador
navegador = webdriver.Chrome()

#passo 1: Pegar a acotação do dólar

#entrar no google
navegador.get("https://google.com.br")

#pesquisar no google a "cotação do dolar"
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação dolar")
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
#pegar a cotação que tá no google
cotacao_dolar = navegador.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

#passo 2: Pegar a cotação do euro
navegador.get("https://google.com.br")
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação euro")
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

#passo 3: Pegar a cotação do ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element("xpath", '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",",".")

navegador.quit()

#passo 4: Atualizar a base de preços (atualizando o preço de compra e o de venda)
import pandas as pd
tabela=pd.read_excel("Produtos.xlsx")
print(tabela)

#atualizar a coluna de cotação
#eu quero editar a coluan cotaçao onde a moeda e = dolar
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

#atualizar a coluna de preço de compra = preço original * cotação
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

#atualizar a coluna de venda = preço de compra * margen
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

print(tabela)

#passo 5: Exportar a base de preços atualizados
tabela.to_excel("Produtos Novo.xlsx", index=False)