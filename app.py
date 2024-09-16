import openpyxl
import pyautogui
workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    #cliente
    pyautogui.click(1741,475,duration=1.5)
    pyautogui.write(linha[0].value)
    #produto
    pyautogui.click(1712,498,duration=1.5)
    pyautogui.write(linha[1].value)
    #quantidade
    pyautogui.click(1731,527,duration=1.5)
    pyautogui.write(str(linha[2].value))
    #categoria  
    pyautogui.click(1806,551,duration=1.5)
    pyautogui.write(linha[3].value)
    #salvar
    pyautogui.click(1648,576,duration=1.5)
    #ok
    pyautogui.click(920,588,duration=1.5)