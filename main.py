import pyautogui
import pyperclip
import time
import pandas as pd
from IPython.display import display

pyautogui.PAUSE = 2

# abra algum navegador e execute este código para iniciar o processo

# altera a janela para o navegador
pyautogui.hotkey('alt', 'tab')

# abre uma nova aba
pyautogui.hotkey('ctrl', 't')

# acessa o drive para fazer o download do arquivo
link = "https://drive.google.com/drive/u/1/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(15)

#posição da pasta no google drive
pyautogui.click(x=388, y=376, clicks=2)
time.sleep(15)

#posição do arquivo
pyautogui.click(x=388, y=376)

#posição do menu para fazer o download
pyautogui.click(x=1157, y=182)

# clica em download
pyautogui.click(x=916, y=583)

tabela = pd.read_excel(r".\Vendas - Dezembro.xlsx")

# mostra a tabela
display(tabela)

# soma da coluna 'valor final' e 'quantidade' da tabela
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento, quantidade)

# automação do envio de e-mail
pyautogui.hotkey("alt","tab")
pyautogui.hotkey("ctrl", "t")
link = "https://mail.google.com/"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#posições do botão "escrever" novo email
time.sleep(10)
pyautogui.click(x=98, y=198)
time.sleep(25)

#email do destinatário
pyautogui.write("laissi.vetech+teste@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")

# campo assunto
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

# corpo do email
texto = f"""
Prezados, bom dia

O faturamento foi de R$ {faturamento:,.2f}
A quantidade foi de {quantidade:,}

Att.
Vetech
"""
pyautogui.write(texto)

# enviar email
pyautogui.hotkey("ctrl", "enter")




