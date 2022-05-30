from distutils.cmd import Command
import tkinter as tk
from sqlite3 import Row
from xml.etree.ElementPath import findtext
import pandas as pd
import win32com.client as win32
from tkinter import messagebox
import sys
import os
from pathlib import Path
import time


def validar_usuario():
    login = str(campo_1.get())
    senha = str(campo_2.get())

    if login == "zaire" and senha == "asdf":
        responsavel = "Zaire Landy"
        escolhe_ECO(responsavel)

    elif login == "vinicius" and senha == "asdf":
        responsavel = "Vinicius Lopes"
        escolhe_ECO(responsavel)

    elif login == "favrin" and senha == "asdf":
        responsavel = "Guilherme Favrin"
        escolhe_ECO(responsavel)

    elif login == "tiago" and senha == "asdf":
        responsavel = "Tiago Simões"
        escolhe_ECO(responsavel)

    else:
        label_3 = tk.Label(janela, text="Login e/ou login incorretos!")
        label_3.grid(row=2, column=1)


def escolhe_ECO(responsavel):

    janela_2 = tk.Toplevel(janela)
    janela_2.title("BOM")
    janela_2.geometry("330x150")
    janela_2.resizable(width=False, height=False)
    label_4 = tk.Label(janela_2, text="Ano:   ", font=("calibri", 11, "bold"))
    label_5 = tk.Label(janela_2, text="Número da ECO:   ",
                       font=("calibri", 11, "bold"))
    label_6 = tk.Label(janela_2, text=" ")
    campo_3 = tk.Entry(janela_2)
    campo_4 = tk.Entry(janela_2)
    # o lambda é necessário para passar funções com argumentos
    botao_2 = tk.Button(janela_2, text="Disponibilizar", command=lambda: disponibiliza(
        str(campo_3.get()), str(campo_4.get()), responsavel))

    label_4.grid(row=0, column=0)
    label_5.grid(row=1, column=0)
    label_6.grid(row=2, column=0)
    campo_3.grid(row=0, column=1)
    campo_4.grid(row=1, column=1)
    botao_2.grid(row=3, column=1)


def disponibiliza(ano, numero, responsavel):

    janela.destroy()
    # CAMINHO DAS BOMs
    path_analiseimpactoo = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\ECO\AnaliseImpacto'
    path_analiseimpactoo = path_analiseimpactoo + "\ECOs_ " + \
        ano + '\ECO-' + ano + '-' + '00' + numero + r'\atualizado'
    # DESTINO DOS HTMLs DAS BOMs
    path_html = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Web\htm'
    # DESTINO DOS xlsx DAS BOMs
    path_xls = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Web\xls'
    # FTP
    path_ftp = r'ftp://ftp1.padtec.com.br/Placas/Hardware/BOM/'

    i = 0
    BOMs = ""
    lista_atualizada = os.listdir(
        path_analiseimpactoo)
    print(lista_atualizada)

    while i < len(lista_atualizada):

        if Path(lista_atualizada[i]).suffix == '.xls':

            xls = lista_atualizada[i]

            if(xls[:5]) == "6.800":

                BOM = xls
                BOMs = BOMs + xls[:-4] + "; "

                destino_1 = path_xls + "\\" + BOM
                destino_2 = path_html + "\\" + BOM[:-4] + '.html'
                destino_3 = path_ftp + BOM

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                testando = path_analiseimpactoo + "\\" + lista_atualizada[i]
                print(testando)
                wb = excel.Workbooks.Open(testando)

                j = 1
                lista = []
                while j < wb.Sheets.Count:
                    versao = wb.Sheets(j).Name
                    if versao[:1] == 'V':
                        lista.append(int(versao[1:]))
                    j = j + 1
                versao = 'V' + str(max(lista))

                print(destino_1)
                print(destino_2)
                print(destino_3)

                excel.DisplayAlerts = False
                ws = wb.Worksheets(versao)
                Range = ws.Range("A:A")
                ws.Cells(Range.Find("Status:").Row, Range.Find(
                    "Status:").Column + 1).Value = "Disponível na Web"
                ws.Cells(Range.Find("Status:").Row + 1,
                         Range.Find("Status:").Column + 1).Value = responsavel

                wb.SaveAs(destino_1, FileFormat=51)  # xlsx
                wb.SaveAs(destino_2, FileFormat=44)  # html
                wb.SaveAs(destino_1, FileFormat=51)  # xls

                wb.Close()
                excel.Application.Quit()
                versao = "nada"
            # else:
            #     messagebox.showerror(message="Os códigos não estão corretos.", title="BOMs")
            #     sys.exit(0)
        i = i + 1
    print(BOMs)
    messagebox.showinfo(
        message="As seguintes BOMs foram disponibilizadas: " + BOMs, title="BOMs")
    sys.exit(0)

    # for quantidade in range(0, len(BOM.index)):
    # if len(str(BOM.iat[quantidade, 0])) == 12:


janela = tk.Tk()
janela.title("Login")
janela.geometry("230x100")
janela.resizable(width=False, height=False)
label_1 = tk.Label(janela, text="Login:   ", font=("calibri", 11, "bold"))
label_2 = tk.Label(janela, text="Senha:   ", font=("calibri", 11, "bold"))
label_3 = tk.Label(janela, text=" ")
campo_1 = tk.Entry(janela)
campo_2 = tk.Entry(janela, show="*")
botao_1 = tk.Button(text="Acessar", command=validar_usuario)
label_1.grid(row=0, column=0)
label_2.grid(row=1, column=0)
label_3.grid(row=2, column=1)
campo_1.grid(row=0, column=1)
campo_2.grid(row=1, column=1)
botao_1.grid(row=3, column=1)

janela.mainloop()
