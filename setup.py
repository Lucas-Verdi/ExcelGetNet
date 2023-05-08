from cx_Freeze import setup, Executable

setup(
    name="Excel GetNet",
    version="2.0",
    description='''Excel GetNet
Autor: Lucas Arnaut Verdi
Versão: 2.0
Data: 27/04/2023''',
    executables=[Executable("excelgetnet.py", base="Win32GUI")],
)

#cxfreeze SeparadorDePedidos.py --target-dir Separador-Separador1.2

#python setup.py build