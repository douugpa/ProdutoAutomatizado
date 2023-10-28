import openpyxl
import pyautogui
import pyperclip

# Abrir a planilha 
workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
sheet_vendas = workbook['vendas']

# Ler o valor dos campos
# Inserir os valores no sistema
for linha in sheet_vendas.iter_rows(min_row=2):
   
   # Nome   
   pyautogui.click(651,297, duration=0.2)
   nome = linha[0].value
   pyperclip.copy(nome)
   pyautogui.hotkey('ctrl', 'v')
   pyautogui.hotkey('tab')

   # Produto
   produto = linha[1].value
   pyperclip.copy(produto)
   pyautogui.hotkey('ctrl', 'v')
   pyautogui.hotkey('tab')

   # Quantidade
   quantidade = linha[2].value
   pyperclip.copy(quantidade)
   pyautogui.hotkey('ctrl', 'v')
   pyautogui.hotkey('tab')

   # Categoria
   categoria = linha[3].value
   pyperclip.copy(categoria)
   pyautogui.hotkey('ctrl', 'v')
   pyautogui.hotkey('tab') 

   pyautogui.click(596,403, duration=0.1)  # Click no Salvar 
   pyautogui.hotkey('tab') 
   pyautogui.press('enter') # Click na Confirmação
   pyautogui.hotkey('tab') 
   

