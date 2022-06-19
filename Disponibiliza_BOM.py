import ftplib
from logging import shutdown
import time
from pathlib import Path
import os
import sys
from tkinter import messagebox
import win32com.client as win32
import pandas as pd
from xml.etree.ElementPath import findtext
from turtle import clear
from sqlite3 import Row
import tkinter as tk
from distutils.cmd import Command
import shutil

ftp = ftplib.FTP("ftp1.padtec.com.br")
ftp.login("placas_hardware", "g5hyYd")

path_htm = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Web\htm'
path_xls = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Web\xls'
path_bom = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\BOM'
path_ep = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Produto\EPs'


def validar_usuario():
    login = str(campo_1.get())
    senha = str(campo_2.get())

    if login == "zaire" and senha == "asdf":
        responsavel = "Zaire Couto"
        janela.destroy()
        escolhe_ECO(responsavel)

    elif login == "vinicius" and senha == "asdf":
        responsavel = "Vinicius Lopes"
        janela.destroy()
        escolhe_ECO(responsavel)

    elif login == "favrin" and senha == "asdf":
        responsavel = "Guilherme Favrin"
        janela.destroy()
        escolhe_ECO(responsavel)

    elif login == "tiago" and senha == "asdf":
        responsavel = "Tiago Simões"
        janela.destroy()
        escolhe_ECO(responsavel)

    else:
        label_3 = tk.Label(janela, text="Login e/ou login incorretos")
        label_3.grid(row=2, column=1)


def escolhe_ECO(responsavel):

    janela2 = tk.Tk()
    janela2.title("BOM")
    janela2.geometry("330x150")
    janela2.resizable(width=False, height=False)
    label_4 = tk.Label(janela2, text="Ano:   ", font=("calibri", 11, "bold"))
    label_5 = tk.Label(janela2, text="Número da ECO:   ",
                       font=("calibri", 11, "bold"))
    label_6 = tk.Label(janela2, text=" ")
    campo_3 = tk.Entry(janela2)
    campo_4 = tk.Entry(janela2)
    # o lambda é necessário para passar funções com argumentos
    botao_2 = tk.Button(janela2, text="Disponibilizar", command=lambda: teste(
        str(campo_3.get()), str(campo_4.get()), responsavel, janela2, chkValue.get()))
    chkValue = tk.BooleanVar()
    chkValue.set(False)
    botaocheck = tk.Checkbutton(
        janela2, text='Disponibilizar EPs.', var=chkValue)

    label_4.grid(row=0, column=0)
    label_5.grid(row=1, column=0)
    label_6.grid(row=3, column=0)
    campo_3.grid(row=0, column=1)
    campo_4.grid(row=1, column=1)
    botao_2.grid(row=4, column=1)
    botaocheck.grid(row=2, column=1)


def teste(ano, numero, responsavel, janela2, eps):

    path_analiseimpacto = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\ECO\AnaliseImpacto'
    path_analiseimpacto = path_analiseimpacto + "\ECOs_ " + \
        ano + '\ECO-' + ano + '-' + '00' + numero + r'\atualizado'

    try:
        folder = os.listdir(path_analiseimpacto)

    except:
        messagebox.showerror(
            message="Ano e/ou número da BOM incorretos!: ", title="ERRO")
    else:
        disponibiliza(ano, numero, responsavel, janela2, eps)


def disponibiliza(ano, numero, responsavel, janela2, eps):

    janela2.destroy()
    global path_htm
    global path_xls
    global path_ep

    path_analiseimpacto = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\ECO\AnaliseImpacto'
    path_analiseimpacto = path_analiseimpacto + "\ECOs_ " + \
        ano + '\ECO-' + ano + '-' + '00' + numero + r'\atualizado'

    BOMs = ""
    folder = os.listdir(
        path_analiseimpacto)

    start = time.time()

    ep = []

    for i in folder:

        if Path(i).suffix == '.xls':

            xls = i

            if(xls[:5]) == "6.800":

                BOM = xls
                BOMs = BOMs + xls[:-4] + "; "

                destino_1 = path_bom + "\\" + BOM
                destino_2 = path_xls + "\\" + BOM
                destino_3 = path_htm + "\\" + BOM[:-4] + '.htm'

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                arquivo = path_analiseimpacto + "\\" + i
                wb = excel.Workbooks.Open(arquivo)

                ftp.cwd(r"/Placas/Hardware/BOM/")
                with open(arquivo, "rb") as file:
                    ftp.storbinary(f"STOR {i}", file)

                j = 1
                lista = []

                while j < wb.Sheets.Count:
                    versao = wb.Sheets(j).Name
                    if versao[:1] == 'V':
                        lista.append(int(versao[1:]))
                    j = j + 1
                versao = 'V' + str(max(lista))

                excel.DisplayAlerts = False
                ws = wb.Worksheets(versao)
                Range = ws.Range("A:A")
                ws.Cells(Range.Find("Status:").Row, Range.Find(
                    "Status:").Column + 1).Value = "Disponível em Web"
                ws.Cells(Range.Find("Status:").Row + 1,
                         Range.Find("Status:").Column + 1).Value = responsavel

                wb.SaveAs(destino_1, FileFormat=56)  # xls
                wb.SaveAs(destino_2, FileFormat=56)  # xls
                wb.SaveAs(destino_3, FileFormat=44)  # htm

                wb.Close()
                excel.Application.Quit()
                versao = "nada"

        elif eps is True:

            if Path(i).suffix == '.xlsx':

                xlsx = i

                if (xlsx[:5]) != "6.800":

                    shutil.copy((path_analiseimpacto + "\\" + xlsx),
                                (path_ep + "\\" + xlsx))

    end = time.time()

    if BOMs == "":

        messagebox.showwarning(
            message=f"Nenhuma BOM foi encontrada no local indicado.", title="BOMs")

    else:

        messagebox.showinfo(
            message=f"TEMPO DE EXECUÇÃO: {int((end - start)//60)} minuto(s) e {int((end - start)%60)} segundos. As seguintes BOMs foram disponibilizadas: " + BOMs, title="BOMs")
    ftp.quit()
    sys.exit(0)


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
