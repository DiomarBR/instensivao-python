import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# pyautogui.click -> clicar
# pyautogui.press -> apertar 1 tecla
# pyautogui.hotkey -> conjunto de teclas
# pyautogui.write -> escreve um texto

# Passo 1: Entrar no sistema da empresa (no nosso caso é o link do drive)
pyautogui.click(x=496, y=747)
time.sleep(5)
pyperclip.copy("Chrome")
pyautogui.hotkey("ctrl","V")
pyautogui.press("enter")

pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar no sistema e encontrar a base de vendas (entrar na pasta exportar)
pyautogui.click(x=330, y=287, clicks=2)
time.sleep(6)

# Passo 3: Fazer o download da base de vendas
pyautogui.click(x=390, y=279) # clicar no arquivo
pyautogui.click(x=1154, y=185) # clicar nos 3 pontinhos
pyautogui.click(x=934, y=626) # clicar no fazer download
time.sleep(5) # esperar o download acabar



# Passo 4: Importar a base de vendas pro Python

tabela = pd.read_excel(r"C:\Users\Diomar Gonçalves\Downloads\Vendas - Dez.xlsx")
print(tabela)

# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)

# Passo 6: Enviar um e-mail para a diretoria com os indicadores de venda

# abrir aba
pyautogui.hotkey("ctrl", "t")

# entrar no link do email - https://mail.google.com/mail/u/0/#inbox
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# clicar no botão escrever
pyautogui.click(x=78, y=192)
time.sleep(5)
# preencher as informações do e-mail
pyautogui.write("lissamely03@gmail.com")
pyautogui.press("tab") # selecionar o email

pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab") # pular para o campo de corpo do email

texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Lira do Python
"""

# formatação dos números (moeda, dinheiro)

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")