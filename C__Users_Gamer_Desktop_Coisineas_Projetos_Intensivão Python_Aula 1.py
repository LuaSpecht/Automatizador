#!/usr/bin/env python
# coding: utf-8

# In[35]:


#!pip install pyautogui


# In[36]:


import pyautogui #biblioteca que "acessa" o seu mouse, teclado e tela 
import pyperclip #biblioteca que copia e cola
import time #biblioteca que manipula o tempo


pyautogui.PAUSE = 1

# Passo 1: Entrar no Sistema (no caso é entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v") 
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar no sistema e encontrar a base de dados
pyautogui.click(x=385, y=306, clicks=2)

time.sleep(2)

# Passo 3: Fazer Download da base de dados
pyautogui.click(x=373, y=389) 
pyautogui.click(x=1161, y=192)
pyautogui.click(x=927, y=590)

time.sleep(7)


# In[39]:


# Passo 4: Calcular indicadores (faturamento da empresa, quantidade de produtos)
import pandas as pd #biblioteca que entende tabelas

tabela = pandas.read_excel(r"C:\Users\Gamer\Downloads\Vendas - Dez.xlsx")

quantidade = tabela["Quantidade"].sum()

faturamento = tabela["Valor Final"].sum()


# In[38]:


# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 6: Mandar um e-mail para a diretoria com os indicadores

#abre para escrever o e-mail
pyautogui.click(x=84, y=205)

time.sleep(2)

#seleciona o destinatário do e-mail
pyautogui.write("luaspecht@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")

#imprime o assunto do e-mail
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

#corpo do e-mail
texto = f"""
Ola, bom dia!

O faturamento é de: R${faturamento:,.2f}
O quantidade de problemas: {quantidade:,}
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")


#envia o e-mail
pyautogui.hotkey("ctrl", "enter")

