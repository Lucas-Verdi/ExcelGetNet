import tkinter as tk
from tkinter import *
from tkinter import filedialog
import xlwings
import pyautogui
import win32com.client as win32
import pythoncom
import pywintypes
import ctypes


def excelgetnet():
    #Criando janela para selecionar o arquivo
    root = tk.Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename()

    #Abrindo a planilha selecionada
    pastadetrabalho = xlwings.Book(arquivo)

    #Abre o Excel em tela cheia
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.Visible = True
    xl.Workbooks.Open(arquivo)
    xl.ActiveWindow.WindowState = win32.constants.xlMaximized

    # Localiza a planilha com nome "Sintético"
    sintetico_sheet = None
    for sheet in pastadetrabalho.sheets:
        if sheet.name == "Sintético":
            sintetico_sheet = sheet
            break

    # Verifica se a planilha foi encontrada
    if sintetico_sheet is not None:
        # Move o mouse para a célula A1 da planilha "Sintético"
        sintetico_sheet.api.Activate()
        #pyautogui.moveTo(sintetico_sheet.range('A1').left, sintetico_sheet.range('A1').top)
    else:
        print("Planilha 'Sintético' não encontrada")

    #Selecionando a planilha
    planilha = pastadetrabalho.sheets[1]

    #Achando a palavra total e movendo o mouse até ela
    celula = planilha.api.Cells.Find('Total')
    if celula is not None:
        print('Célula encontrada:', celula.Address)
        celula.Select()
        #pyautogui.PAUSE = 0.5
        #pyautogui.moveTo(celula.left, celula.top, duration=0.25)
    else:
        print('Erro: célula não encontrada')

    #Adicionando a fórmula PROCV
    pyautogui.sleep(0.5)
    pyautogui.moveTo(400, 0)
    pyautogui.click()
    pyautogui.sleep(0.5)
    pyautogui.press('right')
    pyautogui.press('right')
    pyautogui.typewrite('=PROCV(B12;B:B;1;0)')
    pyautogui.press('ENTER')

    #Adicionando o filtro para total
    intervalo = planilha.range('C1:C' + str(planilha.cells.last_cell.row)).api
    intervalo.AutoFilter(1, "Total")
    pyautogui.press('up')

    #Arrastando a formula até o fim da coluna
    pyautogui.hotkey('ctrl', 'c')
    col = 'E'
    last_row = planilha.cells.last_cell.row
    define_row = 850
    range_string = f'{col}5:{col}{define_row}'
    planilha.range(range_string).select()
    pyautogui.hotkey('ctrl', 'v')





#INTERFACE
janela = Tk()
janela.title('ExcelGetnet')
Label1 = Label(janela, text='Selecione um modelo de formatação:')
Label1.grid(column=0, row=0, padx=10, pady=10)
Botao1 = Button(janela, text='GetNet')
Botao1.bind("<Button>",  lambda e: excelgetnet())
Botao1.grid(column=0, row=1, padx=10, pady=10)
janela.mainloop()


