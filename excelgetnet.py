import tkinter as tk
from tkinter import *
from tkinter import filedialog
import xlwings
import pyautogui
import win32com.client as win32
import pythoncom
import pywintypes
import ctypes
from threading import Thread

datas = []
cont = 4

class Th(Thread):

    def __init__(self, num):
        Thread.__init__(self)
        self.num = num
    def run(self):

        global datas
        global cont

        #Criando janela para selecionar o arquivo
        root = tk.Tk()
        root.withdraw()
        arquivo = filedialog.askopenfilename()

        #Abrindo a planilha selecionada
        pastadetrabalho = xlwings.Book(arquivo)

        # Abre o Excel em tela cheia
        excel_window = pyautogui.getWindowsWithTitle("Excel")[0]
        excel_window.maximize()

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
        else:
            print('Erro: célula não encontrada')

        last_row = planilha.range('B4').end('down').row
        for i in range(4, last_row + 1):
            data_local = planilha.range('B{}'.format(i)).value
            datas.append(data_local)
            cont += 1

        #Adicionando a fórmula PROCV
        pyautogui.sleep(0.5)
        pyautogui.moveTo(400, 0)
        pyautogui.click()
        pyautogui.sleep(0.5)
        pyautogui.press('right')
        pyautogui.press('right')
        pyautogui.typewrite('=PROCV(B{};B:B;1;0)'.format(cont - 1))
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


def start():
    a = Th(1)
    a.start()


#INTERFACE
janela = Tk()
janela.title('Getnet')
Label1 = Label(janela, text='Selecione um modelo de formatação:')
Label1.grid(column=0, row=0, padx=10, pady=10)
Botao1 = Button(janela, text='GetNet')
Botao1.bind("<Button>",  lambda e: start())
Botao1.grid(column=0, row=1, padx=10, pady=10)
janela.mainloop()


