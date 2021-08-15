import pyautogui
import pyperclip
import time
import pandas as pd
from IPython.display import display

pyautogui.PAUSE = 2

# pyautogui.hotkey('alt', 'tab')
# pyautogui.hotkey('ctrl', 't')
# link = "https://drive.google.com/drive/u/1/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
# pyperclip.copy(link)
# pyautogui.hotkey("ctrl","v")
# pyautogui.press("enter")


# time.sleep(14)
# pyautogui.click(x=388, y=376, clicks=2) #posição da pasta no google drive
# time.sleep(15)
# pyautogui.click(x=388, y=376) #posição do arquivo
# pyautogui.click(x=1157, y=182) #posição do menu para fazer o download
# pyautogui.click(x=916, y=583) #download

#'r' para que o python não leia \seguido de letra, uma caractere especial
tabela = pd.read_excel(r"C:\Users\laiss\Desktop\_programacao\python\minicuroso_python\Vendas - Dezembro.xlsx")

# mostra a tabela
display(tabela)


faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
# print(faturamento, quantidade)

#automação do envio de e-mail
# pyautogui.hotkey('alt','tab')
# pyautogui.hotkey('ctrl', 't')
# link = "https://mail.google.com/"
# pyperclip.copy(link)
# pyautogui.hotkey('ctrl', 'v')
# pyautogui.press('enter')


# PEGAR AS POSIÇÕES X e Y
# time.sleep(5)
# print('Posicione o MOUSE')
# time.sleep(10)
# x, y = pyautogui.position()
# print ("Posicao atual do mouse:")
# print ("x = "+str(x)+" y = "+str(y))

