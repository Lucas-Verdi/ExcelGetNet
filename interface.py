import tkinter as tk
from tkinter import *
import excelgetnet
from excelgetnet import *

janela = Tk()
janela.title('ExcelGetnet')
Label1 = Label(janela, text='Selecione um modelo de formatação:')
Label1.grid(column=0, row=0, padx=10, pady=10)
Botao1 = Button(janela, text='GetNet')
Botao1.bind("<Button>",  lambda e: excelgetnet())
Botao1.grid(column=0, row=1, padx=10, pady=10)
janela.mainloop()