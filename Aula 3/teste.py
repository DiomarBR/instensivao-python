from selenium import webdriver #criar o navegador
from selenium.webdriver.common.keys import Keys#permite escrever no navegador
from selenium.webdriver.common.by import By#permite você selecionar itens no navegador
#abrir o navegador
navegador = webdriver.Chrome(executable_path='./chromedriver.exe')

#passo 1: Pegar a acotação do dólar

#entrar no google
navegador.get("http://127.0.0.1:5500/teste%20html/teste.html")

#pesquisar no google a "cotação do dolar"
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação dolar")
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
#pegar a cotação que tá no google
cotacao_dolar = navegador.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

#passo 2: Pegar a cotação do euro
navegador.get("https://google.com.br")
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação euro")
navegador.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)
#passo 3: Pegar a cotação do ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element("xpath", '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",",".")
print(cotacao_ouro)

navegador.quit()