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
from pyautogui import sleep


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

        last_row = planilha.range('D5').end('down').row

        for i in range(4, last_row + 1):
            cell = planilha.range('B{}'.format(i)).value
            print(cell)
            if cell == None:
                pyautogui.moveTo(400, 0)
                pyautogui.click()
                planilha.range('B{}'.format(i - 1)).select()
                sleep(0.1)
                pyautogui.hotkey('ctrl', 'c')
                pyautogui.press('right', presses=3)
                pyautogui.press('down')
                pyautogui.hotkey('ctrl', 'v')

        # Adicionando o filtro para total
        intervalo = planilha.range('C1:C' + str(planilha.cells.last_cell.row)).api
        intervalo.AutoFilter(1, "Total")
        pyautogui.press('up')



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


