#!/usr/bin/env python
# coding: utf-8

# In[50]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

pyautogui.hotkey ('ctrl','t')
pyperclip.copy ('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey ('ctrl', 'v')
pyautogui.press ('enter')
time.sleep (6)
pyautogui.click (x=381, y=279)
pyautogui.press ('enter')
time.sleep (6)
pyautogui.click (x=314, y=438)
time.sleep (6)
pyautogui.click (x=1142, y=167)
time.sleep(6)
pyautogui.click (x=965, y=536)


# In[51]:


import pandas

tabela = pandas.read_excel(r'C:\Users\paula\Downloads\Vendas - Dez.xlsx')
display (tabela)


# In[46]:


quantidade = tabela ['Quantidade'].sum()
faturamento = tabela ['Valor Final'].sum()
print (quantidade)
print (faturamento)


# In[48]:


quantidade = tabela ['Quantidade'].sum()
faturamento = tabela ['Valor Final'].sum()

pyautogui.hotkey ('ctrl','t')
pyperclip.copy ('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey ('ctrl', 'v')
pyautogui.press ('enter')
time.sleep (15)
pyautogui.click (x=88, y=187)
time.sleep (15)

pyautogui.write ('paula.rsb@hotmail.com')
pyautogui.press ('tab')
pyautogui.press ('tab')

pyperclip.copy ('Relatório de Vendas')
pyautogui.hotkey ('ctrl', 'v')
pyautogui.press ('tab')

texto = f"""
Prezados, boa tarde. Espero que estejam bem.
Segue neste e-mail, o faturamento do dia de hoje.
Quantidade:{quantidade:,}
Faturamento: R${faturamento:,.2f}

Tenham um ótimo dia, qualquer dúvida estou à disposição.
     
Att
Paula Barros
"""

pyperclip.copy (texto)
pyautogui.hotkey ('ctrl','v')
time.sleep (15)
pyautogui.hotkey ('ctrl', 'enter')


# In[ ]:





# In[34]:


time.sleep(8)
pyautogui.position()


# In[ ]:




