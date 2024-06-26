#! /usr/bin/env python3
#  -*- coding: utf-8 -*-
#
# GUI module generated by PAGE version 8.0
#  in conjunction with Tcl version 8.6
#    Mar 30, 2024 08:33:17 PM -03  platform: Windows NT

import sys
import numpy as np
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk
from ctk_maskedentry import CTkMaskedEntry, Mask
from tkinter.constants import *
from tkinter import filedialog
import os.path
import math
from datetime import datetime

from PlaneelhaOutputer import  PlaneelhaOutputer
from PlaneelhaOutputer import load_data

_location = os.path.dirname(__file__)

_bgcolor = '#d9d9d9'
_fgcolor = '#000000'
_tabfg1 = 'black' 
_tabfg2 = 'white' 
_bgmode = 'light' 
_tabbg1 = '#d9d9d9' 
_tabbg2 = 'gray40' 
_lotesSpan = 6

def is_valid_datetime(date_string, date_format):
    try:
        datetime.strptime(date_string, date_format)
        return True
    except ValueError:
        return False

class TopLevelFormulario:

    def removeAllLotes(self):
        for frame in self.FramesLote:
            frame.grid_forget()
        self.FramesLote = []
        self.LabelsLote = []
        self.EntriesLote = []
        self.lotesQtd = []

    def agrupamento_radio_change(self):
        if (self.agrupamento.get() == 0):
            self.ButtonAtualizar.configure(state='disabled')
            self.LabelframeQtd.configure(text='''Menor preço por Item''')
            self.LabelQtd.configure(text='''Qtd. de itens:''')
            self.removeAllLotes()
        else:
            self.ButtonAtualizar.configure(state='normal')
            self.LabelframeQtd.configure(text='''Menor preço por Lote''')
            self.LabelQtd.configure(text='''Qtd. de lotes:''')

    def number_mask(self, P):
        if str.isdigit(P) or P == "":
            return True
        else:
            return False

    def button_atualizar_callback(self):
        qtd = 0
        try:
            qtd = int(self.qtd.get())
        except:
            qtd = 0
        if qtd > 0:
            vcmd = self.CanvasLotesQtd.register(self.number_mask)

            self.removeAllLotes()

            for i in range(qtd):
                tVar = tk.IntVar()
                self.lotesQtd.append(tVar)

                FrameLote = ctk.CTkFrame(master=self.CanvasLotesQtd)
                FrameLote.grid(row=i, column=0, sticky='ew')
                self.FramesLote.append(FrameLote)

                nomeLote = "Lote " + str(i+1) + ":"
                LabelLote = ctk.CTkLabel(master=FrameLote, text=nomeLote)
                LabelLote.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
                LabelLote.configure(font=self.FontLabel)
                self.LabelsLote.append(LabelLote)

                EntryLote = ctk.CTkEntry(master=FrameLote)
                EntryLote.grid(row=0, column=1, padx=5, pady=5, columnspan=_lotesSpan-1, sticky="nsew")
                EntryLote.configure(corner_radius=0)
                EntryLote.configure(textvariable=tVar)
                EntryLote.configure(border_width=1)
                EntryLote.configure(font=self.FontEntry)
                EntryLote.configure(validate='all')
                EntryLote.configure(validatecommand=(vcmd, '%P'))
                EntryLote.delete(0, tk.END)
                self.EntriesLote.append(EntryLote)
                self.CanvasLotesQtd.grid_rowconfigure(i, weight=1)
                
                for i in range(_lotesSpan):
                    FrameLote.grid_columnconfigure(i, weight=1, uniform="foo")
                
                self.CanvasLotesQtd.columnconfigure(0, weight=1)

    def check_form(self):
        if self.orgao.get() == "":
            tk.messagebox.showwarning("Atenção", "Falta o nome do órgão!")
        elif self.codLicitacao.get() == "":
            tk.messagebox.showwarning("Atenção", "Falta o número da licitação!")
        elif self.dataAbertura.get() == "":
            tk.messagebox.showwarning("Atenção", "Falta a data da licitação!")
        elif not is_valid_datetime(self.dataAbertura.get(), "%d/%m/%Y"):
            tk.messagebox.showwarning("Atenção", "Insira uma data válida!")
        elif not is_valid_datetime(self.horaAbertura.get(), "%H:%M"):
            tk.messagebox.showwarning("Atenção", "Insira um horário válido!")
        elif self.agrupamento.get() == 0 and self.qtd.get() == "":
            tk.messagebox.showwarning("Atenção", "Insira a quantidade de itens!")
        elif self.agrupamento.get() == 1 and (self.qtd.get() == "" or not self.lotesQtd):
            tk.messagebox.showwarning("Atenção", "Se estiver trabalhando com lotes, insira a quantidade e atualize a lista!")
        elif self.agrupamento.get() == 1 and int(self.qtd.get()) != len(self.lotesQtd):
            tk.messagebox.showwarning("Atenção", "A quantidade de lotes não bate com o número de campos, atualize a lista!")
        else:
            return True
        return False
    
    def button_criar_callback(self):
        if self.check_form() and self.create_workbook():
            self.exitStatus = True
            self.top.destroy()
    
    def create_workbook(self):
        template_name = "MODELO DE PLANILHA {} - ".format(self.empresa.get())
        if self.fileDir == "":
            self.filePath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=template_name, filetypes=[("Planilha do Excel", "*.xlsx")])
            if (self.filePath):
                print("Selected file:", self.filePath)
                po = PlaneelhaOutputer(self.getFormInfo())
                return True
        else:
            self.filePath = filedialog.asksaveasfilename(initialdir=self.fileDir, initialfile=template_name, defaultextension=".xlsx", filetypes=[("Planilha do Excel", "*.xlsx")])
            if (self.filePath):
                print("Selected file:", self.filePath)
                po = PlaneelhaOutputer(self.getFormInfo())
                return True
        return False

    def tkVars2Integers(self):
        if self.agrupamento.get() == 0: return None
        arr = []
        count = 0
        
        for tkVar in self.lotesQtd:
            try:
                if (count == int(self.qtd.get())): break
            except:
                print("Nothing to do here then.")
                break
            try:
                arr.append(tkVar.get())
            except:
                arr.append(0)
        return arr

    def getFormInfo(self):

        return {
            "orgao": self.orgao.get(),
            "codLicitacao": self.codLicitacao.get(),
            "codProcesso": self.codProcesso.get(),
            "dataAbertura": self.dataAbertura.get(),
            "horaAbertura": self.horaAbertura.get(),
            "empresa": self.empresa.get(),
            "tipo": self.tipo.get(),
            "agrupamento": self.agrupamento.get(),
            "qtd": int(self.qtd.get()),
            "lotesQtd": self.tkVars2Integers(),
            "caminhoArquivo": self.filePath
        }

    def __init__(self, top=None, savedir = ""):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''

        self.FontLabel = ctk.CTkFont(family="Segoe UI", size=12)
        self.FontEntry = ctk.CTkFont(family="Courier New", size=12)

        vcmd = (top.register(self.number_mask))

        top.geometry("645x588+383+93")
        top.minsize(120, 1)
        top.maxsize(3844, 1061)
        top.resizable(1,  1)
        top.title("Iniciar Planilha")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="#000000")
        
        ctk.set_appearance_mode("light")

        self.exitStatus = False
        self.top = top
        self.orgao = tk.StringVar()
        self.codLicitacao = tk.StringVar()
        self.codProcesso = tk.StringVar()
        self.dataAbertura = tk.StringVar()
        self.horaAbertura = tk.StringVar()
        self.empresa = tk.StringVar()
        self.tipo = tk.StringVar()
        self.agrupamento = tk.IntVar()
        self.qtd = tk.StringVar()
        self.lotesQtd = []
        
        self.orgao.set("Prefeitura Municipal de ")

        self.filePath = ""
        self.fileDir = savedir
        
        self.empresa.set("GI")
        self.tipo.set("PREGÃO")
        
        dateMask = Mask('fixed', '99/99/9999')
        timeMask = Mask('fixed', '99:99')
        
        self.Labelframe1 = tk.LabelFrame(self.top)
        self.Labelframe1.place(relx=0.016, rely=0.015, relheight=0.957
                , relwidth=0.972)
        self.Labelframe1.configure(relief='groove')
        self.Labelframe1.configure(font="-family {Segoe UI} -size 9")
        self.Labelframe1.configure(foreground="#000000")
        self.Labelframe1.configure(text='''Novo modelo de planilha''')
        self.Labelframe1.configure(background="#d9d9d9")
        self.Labelframe1.configure(highlightbackground="#d9d9d9")
        self.Labelframe1.configure(highlightcolor="#000000")

        self.Label7 = tk.Label(self.Labelframe1)
        self.Label7.place(relx=0.032, rely=0.721, height=21, width=126
                , bordermode='ignore')
        self.Label7.configure(activebackground="#d9d9d9")
        self.Label7.configure(activeforeground="black")
        self.Label7.configure(anchor='w')
        self.Label7.configure(background="#d9d9d9")
        self.Label7.configure(compound='left')
        self.Label7.configure(disabledforeground="#a3a3a3")
        self.Label7.configure(font="-family {Segoe UI} -size 9")
        self.Label7.configure(foreground="#000000")
        self.Label7.configure(highlightbackground="#d9d9d9")
        self.Label7.configure(highlightcolor="#000000")
        self.Label7.configure(text='''Participar pela:''')

        self.RadioID = tk.Radiobutton(self.Labelframe1)
        self.RadioID.place(relx=0.032, rely=0.792, relheight=0.046
                , relwidth=0.191, bordermode='ignore')
        self.RadioID.configure(activebackground="#d9d9d9")
        self.RadioID.configure(activeforeground="black")
        self.RadioID.configure(anchor='w')
        self.RadioID.configure(background="#d9d9d9")
        self.RadioID.configure(compound='left')
        self.RadioID.configure(disabledforeground="#a3a3a3")
        self.RadioID.configure(font="-family {Segoe UI} -size 9")
        self.RadioID.configure(foreground="#000000")
        self.RadioID.configure(highlightbackground="#d9d9d9")
        self.RadioID.configure(highlightcolor="#000000")
        self.RadioID.configure(justify='left')
        self.RadioID.configure(text='''Info Direct''')
        self.RadioID.configure(value='ID')
        self.RadioID.configure(variable=self.empresa)

        self.RadioEB = tk.Radiobutton(self.Labelframe1)
        self.RadioEB.place(relx=0.032, rely=0.829, relheight=0.044
                , relwidth=0.191, bordermode='ignore')
        self.RadioEB.configure(activebackground="#d9d9d9")
        self.RadioEB.configure(activeforeground="black")
        self.RadioEB.configure(anchor='w')
        self.RadioEB.configure(background="#d9d9d9")
        self.RadioEB.configure(compound='left')
        self.RadioEB.configure(disabledforeground="#a3a3a3")
        self.RadioEB.configure(font="-family {Segoe UI} -size 9")
        self.RadioEB.configure(foreground="#000000")
        self.RadioEB.configure(highlightbackground="#d9d9d9")
        self.RadioEB.configure(highlightcolor="#000000")
        self.RadioEB.configure(justify='left')
        self.RadioEB.configure(text='''Embacom''')
        self.RadioEB.configure(value='EB')
        self.RadioEB.configure(variable=self.empresa)

        self.RadioGC = tk.Radiobutton(self.Labelframe1)
        self.RadioGC.place(relx=0.032, rely=0.865, relheight=0.044
                , relwidth=0.191, bordermode='ignore')
        self.RadioGC.configure(activebackground="#d9d9d9")
        self.RadioGC.configure(activeforeground="black")
        self.RadioGC.configure(anchor='w')
        self.RadioGC.configure(background="#d9d9d9")
        self.RadioGC.configure(compound='left')
        self.RadioGC.configure(disabledforeground="#a3a3a3")
        self.RadioGC.configure(font="-family {Segoe UI} -size 9")
        self.RadioGC.configure(foreground="#000000")
        self.RadioGC.configure(highlightbackground="#d9d9d9")
        self.RadioGC.configure(highlightcolor="#000000")
        self.RadioGC.configure(justify='left')
        self.RadioGC.configure(text='''Gold Comércio''')
        self.RadioGC.configure(value='GC')
        self.RadioGC.configure(variable=self.empresa)

        self.RadioGI = tk.Radiobutton(self.Labelframe1)
        self.RadioGI.place(relx=0.032, rely=0.757, relheight=0.046, relwidth=0.19
                , bordermode='ignore')
        self.RadioGI.configure(activebackground="#d9d9d9")
        self.RadioGI.configure(activeforeground="black")
        self.RadioGI.configure(anchor='w')
        self.RadioGI.configure(background="#d9d9d9")
        self.RadioGI.configure(compound='left')
        self.RadioGI.configure(disabledforeground="#a3a3a3")
        self.RadioGI.configure(font="-family {Segoe UI} -size 9")
        self.RadioGI.configure(foreground="#000000")
        self.RadioGI.configure(highlightbackground="#d9d9d9")
        self.RadioGI.configure(highlightcolor="#000000")
        self.RadioGI.configure(justify='left')
        self.RadioGI.configure(text='''Gráfica Iguaçu''')
        self.RadioGI.configure(value='GI')
        self.RadioGI.configure(variable=self.empresa)

        self.RadioDispensa = tk.Radiobutton(self.Labelframe1)
        self.RadioDispensa.place(relx=0.032, rely=0.504, relheight=0.046
                , relwidth=0.191, bordermode='ignore')
        self.RadioDispensa.configure(activebackground="#d9d9d9")
        self.RadioDispensa.configure(activeforeground="black")
        self.RadioDispensa.configure(anchor='w')
        self.RadioDispensa.configure(background="#d9d9d9")
        self.RadioDispensa.configure(compound='left')
        self.RadioDispensa.configure(disabledforeground="#a3a3a3")
        self.RadioDispensa.configure(font="-family {Segoe UI} -size 9")
        self.RadioDispensa.configure(foreground="#000000")
        self.RadioDispensa.configure(highlightbackground="#d9d9d9")
        self.RadioDispensa.configure(highlightcolor="#000000")
        self.RadioDispensa.configure(justify='left')
        self.RadioDispensa.configure(text='''Dispensa''')
        self.RadioDispensa.configure(value='DISPENSA')
        self.RadioDispensa.configure(variable=self.tipo)

        self.RadioCotacao = tk.Radiobutton(self.Labelframe1)
        self.RadioCotacao.place(relx=0.032, rely=0.542, relheight=0.044
                , relwidth=0.191, bordermode='ignore')
        self.RadioCotacao.configure(activebackground="#d9d9d9")
        self.RadioCotacao.configure(activeforeground="black")
        self.RadioCotacao.configure(anchor='w')
        self.RadioCotacao.configure(background="#d9d9d9")
        self.RadioCotacao.configure(compound='left')
        self.RadioCotacao.configure(disabledforeground="#a3a3a3")
        self.RadioCotacao.configure(font="-family {Segoe UI} -size 9")
        self.RadioCotacao.configure(foreground="#000000")
        self.RadioCotacao.configure(highlightbackground="#d9d9d9")
        self.RadioCotacao.configure(highlightcolor="#000000")
        self.RadioCotacao.configure(justify='left')
        self.RadioCotacao.configure(text='''Cotação''')
        self.RadioCotacao.configure(value='COTAÇÃO')
        self.RadioCotacao.configure(variable=self.tipo)

        self.RadioPregao = tk.Radiobutton(self.Labelframe1)
        self.RadioPregao.place(relx=0.032, rely=0.467, relheight=0.046
                , relwidth=0.19, bordermode='ignore')
        self.RadioPregao.configure(activebackground="#d9d9d9")
        self.RadioPregao.configure(activeforeground="black")
        self.RadioPregao.configure(anchor='w')
        self.RadioPregao.configure(background="#d9d9d9")
        self.RadioPregao.configure(compound='left')
        self.RadioPregao.configure(disabledforeground="#a3a3a3")
        self.RadioPregao.configure(font="-family {Segoe UI} -size 9")
        self.RadioPregao.configure(foreground="#000000")
        self.RadioPregao.configure(highlightbackground="#d9d9d9")
        self.RadioPregao.configure(highlightcolor="#000000")
        self.RadioPregao.configure(justify='left')
        self.RadioPregao.configure(text='''Pregão''')
        self.RadioPregao.configure(value='PREGÃO')
        self.RadioPregao.configure(variable=self.tipo)

        self.RadioPorLote = tk.Radiobutton(self.Labelframe1)
        self.RadioPorLote.place(relx=0.032, rely=0.668, relheight=0.044
                , relwidth=0.191, bordermode='ignore')
        self.RadioPorLote.configure(activebackground="#d9d9d9")
        self.RadioPorLote.configure(activeforeground="black")
        self.RadioPorLote.configure(anchor='w')
        self.RadioPorLote.configure(background="#d9d9d9")
        self.RadioPorLote.configure(compound='left')
        self.RadioPorLote.configure(disabledforeground="#a3a3a3")
        self.RadioPorLote.configure(foreground="#000000")
        self.RadioPorLote.configure(highlightbackground="#d9d9d9")
        self.RadioPorLote.configure(highlightcolor="#000000")
        self.RadioPorLote.configure(justify='left')
        self.RadioPorLote.configure(text='''Por Lote''')
        self.RadioPorLote.configure(value='1')
        self.RadioPorLote.configure(variable=self.agrupamento)
        self.RadioPorLote.configure(command=lambda:self.agrupamento_radio_change())

        self.RadioPorItem = tk.Radiobutton(self.Labelframe1)
        self.RadioPorItem.place(relx=0.032, rely=0.631, relheight=0.046
                , relwidth=0.191, bordermode='ignore')
        self.RadioPorItem.configure(activebackground="#d9d9d9")
        self.RadioPorItem.configure(activeforeground="black")
        self.RadioPorItem.configure(anchor='w')
        self.RadioPorItem.configure(background="#d9d9d9")
        self.RadioPorItem.configure(compound='left')
        self.RadioPorItem.configure(disabledforeground="#a3a3a3")
        self.RadioPorItem.configure(foreground="#000000")
        self.RadioPorItem.configure(highlightbackground="#d9d9d9")
        self.RadioPorItem.configure(highlightcolor="#000000")
        self.RadioPorItem.configure(justify='left')
        self.RadioPorItem.configure(text='''Por Item''')
        self.RadioPorItem.configure(value='0')
        self.RadioPorItem.configure(variable=self.agrupamento)
        self.RadioPorItem.configure(command=lambda:self.agrupamento_radio_change())

        self.ButtonCriar = tk.Button(self.Labelframe1)
        self.ButtonCriar.place(relx=0.032, rely=0.918, height=26
                , relwidth=0.933, bordermode='ignore')
        self.ButtonCriar.configure(activebackground="#d9d9d9")
        self.ButtonCriar.configure(activeforeground="black")
        self.ButtonCriar.configure(background="#d9d9d9")
        self.ButtonCriar.configure(disabledforeground="#a3a3a3")
        self.ButtonCriar.configure(font="-family {Segoe UI} -size 9")
        self.ButtonCriar.configure(foreground="#000000")
        self.ButtonCriar.configure(highlightbackground="#d9d9d9")
        self.ButtonCriar.configure(highlightcolor="#000000")
        self.ButtonCriar.configure(text='''Criar modelo personalizado''')
        self.ButtonCriar.configure(command=lambda:self.button_criar_callback())

        self.Label6 = tk.Label(self.Labelframe1)
        self.Label6.place(relx=0.032, rely=0.595, height=21, width=126
                , bordermode='ignore')
        self.Label6.configure(activebackground="#d9d9d9")
        self.Label6.configure(activeforeground="black")
        self.Label6.configure(anchor='w')
        self.Label6.configure(background="#d9d9d9")
        self.Label6.configure(compound='left')
        self.Label6.configure(disabledforeground="#a3a3a3")
        self.Label6.configure(font="-family {Segoe UI} -size 9")
        self.Label6.configure(foreground="#000000")
        self.Label6.configure(highlightbackground="#d9d9d9")
        self.Label6.configure(highlightcolor="#000000")
        self.Label6.configure(text='''Menor preço:''')

        self.Label1 = tk.Label(self.Labelframe1)
        self.Label1.place(relx=0.032, rely=0.052, height=23, width=579
                , bordermode='ignore')
        self.Label1.configure(activebackground="#d9d9d9")
        self.Label1.configure(activeforeground="black")
        self.Label1.configure(anchor='w')
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(compound='left')
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font="-family {Segoe UI} -size 9")
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(highlightbackground="#d9d9d9")
        self.Label1.configure(highlightcolor="#000000")
        self.Label1.configure(text='''Nome do órgão:''')

        self.EntryNome = tk.Entry(self.Labelframe1)
        self.EntryNome.place(relx=0.032, rely=0.083, height=20, relwidth=0.915
                , bordermode='ignore')
        self.EntryNome.configure(background="white")
        self.EntryNome.configure(disabledforeground="#a3a3a3")
        self.EntryNome.configure(font="-family {Courier New} -size 10")
        self.EntryNome.configure(foreground="#000000")
        self.EntryNome.configure(highlightbackground="#d9d9d9")
        self.EntryNome.configure(highlightcolor="#000000")
        self.EntryNome.configure(insertbackground="#000000")
        self.EntryNome.configure(selectbackground="#d9d9d9")
        self.EntryNome.configure(selectforeground="black")
        self.EntryNome.configure(textvariable=self.orgao)

        self.Label2 = tk.Label(self.Labelframe1)
        self.Label2.place(relx=0.032, rely=0.117, height=24, width=119
                , bordermode='ignore')
        self.Label2.configure(activebackground="#d9d9d9")
        self.Label2.configure(activeforeground="black")
        self.Label2.configure(anchor='w')
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(compound='left')
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(font="-family {Segoe UI} -size 9")
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(highlightbackground="#d9d9d9")
        self.Label2.configure(highlightcolor="#000000")
        self.Label2.configure(text='''Número da licitação:''')

        self.EntryNumPregao = tk.Entry(self.Labelframe1)
        self.EntryNumPregao.place(relx=0.032, rely=0.153, height=20
                , relwidth=0.198, bordermode='ignore')
        self.EntryNumPregao.configure(background="white")
        self.EntryNumPregao.configure(disabledforeground="#a3a3a3")
        self.EntryNumPregao.configure(font="-family {Courier New} -size 10")
        self.EntryNumPregao.configure(foreground="#000000")
        self.EntryNumPregao.configure(highlightbackground="#d9d9d9")
        self.EntryNumPregao.configure(highlightcolor="#000000")
        self.EntryNumPregao.configure(insertbackground="#000000")
        self.EntryNumPregao.configure(selectbackground="#d9d9d9")
        self.EntryNumPregao.configure(selectforeground="black")
        self.EntryNumPregao.configure(textvariable=self.codLicitacao)

        self.Label4 = tk.Label(self.Labelframe1)
        self.Label4.place(relx=0.032, rely=0.187, height=21, width=126
                , bordermode='ignore')
        self.Label4.configure(activebackground="#d9d9d9")
        self.Label4.configure(activeforeground="black")
        self.Label4.configure(anchor='w')
        self.Label4.configure(background="#d9d9d9")
        self.Label4.configure(compound='left')
        self.Label4.configure(disabledforeground="#a3a3a3")
        self.Label4.configure(font="-family {Segoe UI} -size 9")
        self.Label4.configure(foreground="#000000")
        self.Label4.configure(highlightbackground="#d9d9d9")
        self.Label4.configure(highlightcolor="#000000")
        self.Label4.configure(text='''Número do processo:''')

        self.EntryNumProcesso = tk.Entry(self.Labelframe1)
        self.EntryNumProcesso.place(relx=0.032, rely=0.22, height=20
                , relwidth=0.198, bordermode='ignore')
        self.EntryNumProcesso.configure(background="white")
        self.EntryNumProcesso.configure(disabledforeground="#a3a3a3")
        self.EntryNumProcesso.configure(font="-family {Courier New} -size 10")
        self.EntryNumProcesso.configure(foreground="#000000")
        self.EntryNumProcesso.configure(highlightbackground="#d9d9d9")
        self.EntryNumProcesso.configure(highlightcolor="#000000")
        self.EntryNumProcesso.configure(insertbackground="#000000")
        self.EntryNumProcesso.configure(selectbackground="#d9d9d9")
        self.EntryNumProcesso.configure(selectforeground="black")
        self.EntryNumProcesso.configure(textvariable=self.codProcesso)

        self.Label3 = tk.Label(self.Labelframe1)
        self.Label3.place(relx=0.032, rely=0.252, height=21, width=121
                , bordermode='ignore')
        self.Label3.configure(activebackground="#d9d9d9")
        self.Label3.configure(activeforeground="black")
        self.Label3.configure(anchor='w')
        self.Label3.configure(background="#d9d9d9")
        self.Label3.configure(compound='left')
        self.Label3.configure(disabledforeground="#a3a3a3")
        self.Label3.configure(font="-family {Segoe UI} -size 9")
        self.Label3.configure(foreground="#000000")
        self.Label3.configure(highlightbackground="#d9d9d9")
        self.Label3.configure(highlightcolor="#000000")
        self.Label3.configure(text='''Data de abertura:''')

        self.EntryDataAbertura = CTkMaskedEntry(self.Labelframe1, height=20, mask=dateMask)
        self.EntryDataAbertura.place(relx=0.032, rely=0.284
                , relwidth=0.198, bordermode='ignore')
        self.EntryDataAbertura.configure(corner_radius=0)
        self.EntryDataAbertura.configure(border_width=1)
        self.EntryDataAbertura.configure(font=self.FontEntry)
        #self.EntryDataAbertura.configure(background="white")
        #self.EntryDataAbertura.configure(disabledforeground="#a3a3a3")
        #self.EntryDataAbertura.configure(font="TkFixedFont")
        #self.EntryDataAbertura.configure(foreground="#000000")
        #self.EntryDataAbertura.configure(highlightbackground="#d9d9d9")
        #self.EntryDataAbertura.configure(highlightcolor="#000000")
        #self.EntryDataAbertura.configure(insertbackground="#000000")
        #self.EntryDataAbertura.configure(selectbackground="#d9d9d9")
        #self.EntryDataAbertura.configure(selectforeground="black")
        self.EntryDataAbertura.configure(textvariable=self.dataAbertura)

        self.Label8 = tk.Label(self.Labelframe1)
        self.Label8.place(relx=0.032, rely=0.32, height=19, width=126
                , bordermode='ignore')
        self.Label8.configure(activebackground="#d9d9d9")
        self.Label8.configure(activeforeground="black")
        self.Label8.configure(anchor='w')
        self.Label8.configure(background="#d9d9d9")
        self.Label8.configure(compound='left')
        self.Label8.configure(disabledforeground="#a3a3a3")
        self.Label8.configure(foreground="#000000")
        self.Label8.configure(highlightbackground="#d9d9d9")
        self.Label8.configure(highlightcolor="#000000")
        self.Label8.configure(text='''Hora de abertura:''')

        self.EntryHoraAbertura = CTkMaskedEntry(self.Labelframe1, height=20, mask=timeMask)
        self.EntryHoraAbertura.place(relx=0.032, rely=0.355 
                , relwidth=0.198, bordermode='ignore')
        self.EntryHoraAbertura.configure(corner_radius=0)
        self.EntryHoraAbertura.configure(border_width=1)
        self.EntryHoraAbertura.configure(font=self.FontEntry)
        #self.EntryHoraAbertura.configure(background="white")
        #self.EntryHoraAbertura.configure(disabledforeground="#a3a3a3")
        #self.EntryHoraAbertura.configure(font="-family {Courier New} -size 10")
        #self.EntryHoraAbertura.configure(foreground="#000000")
        #self.EntryHoraAbertura.configure(highlightbackground="#d9d9d9")
        #self.EntryHoraAbertura.configure(highlightcolor="#000000")
        #self.EntryHoraAbertura.configure(insertbackground="#000000")
        #self.EntryHoraAbertura.configure(selectbackground="#d9d9d9")
        #self.EntryHoraAbertura.configure(selectforeground="black")
        self.EntryHoraAbertura.configure(textvariable=self.horaAbertura)

        #_style_code()
        self.TSeparator1 = ttk.Separator(self.Labelframe1)
        self.TSeparator1.place(relx=0.022, rely=0.416, relwidth=0.23
                , bordermode='ignore')

        self.Label5 = tk.Label(self.Labelframe1)
        self.Label5.place(relx=0.032, rely=0.432, height=22, width=126
                , bordermode='ignore')
        self.Label5.configure(activebackground="#d9d9d9")
        self.Label5.configure(activeforeground="black")
        self.Label5.configure(anchor='w')
        self.Label5.configure(background="#d9d9d9")
        self.Label5.configure(compound='left')
        self.Label5.configure(disabledforeground="#a3a3a3")
        self.Label5.configure(font="-family {Segoe UI} -size 9")
        self.Label5.configure(foreground="#000000")
        self.Label5.configure(highlightbackground="#d9d9d9")
        self.Label5.configure(highlightcolor="#000000")
        self.Label5.configure(text='''Tipo de licitação:''')

        self.LabelframeQtd = tk.LabelFrame(self.Labelframe1)
        self.LabelframeQtd.place(relx=0.276, rely=0.13, relheight=0.78
                , relwidth=0.681, bordermode='ignore')
        self.LabelframeQtd.configure(relief='groove')
        self.LabelframeQtd.configure(font="-family {Segoe UI} -size 9")
        self.LabelframeQtd.configure(foreground="#000000")
        self.LabelframeQtd.configure(text='''Menor preço por Item''')
        self.LabelframeQtd.configure(background="#d9d9d9")
        self.LabelframeQtd.configure(highlightbackground="#d9d9d9")
        self.LabelframeQtd.configure(highlightcolor="#000000")

        self.LabelQtd = tk.Label(self.LabelframeQtd)
        self.LabelQtd.place(relx=0.047, rely=0.1, height=24, width=381
                , bordermode='ignore')
        self.LabelQtd.configure(activebackground="#d9d9d9")
        self.LabelQtd.configure(activeforeground="black")
        self.LabelQtd.configure(anchor='w')
        self.LabelQtd.configure(background="#d9d9d9")
        self.LabelQtd.configure(compound='left')
        self.LabelQtd.configure(disabledforeground="#a3a3a3")
        self.LabelQtd.configure(font="-family {Segoe UI} -size 9")
        self.LabelQtd.configure(foreground="#000000")
        self.LabelQtd.configure(highlightbackground="#d9d9d9")
        self.LabelQtd.configure(highlightcolor="#000000")
        self.LabelQtd.configure(text='''Qtd. de itens:''')

        self.EntryQtd = tk.Entry(self.LabelframeQtd)
        self.EntryQtd.place(relx=0.047, rely=0.153, height=20, relwidth=0.628
                , bordermode='ignore')
        self.EntryQtd.configure(background="white")
        self.EntryQtd.configure(disabledforeground="#a3a3a3")
        self.EntryQtd.configure(font="-family {Courier New} -size 10")
        self.EntryQtd.configure(foreground="#000000")
        self.EntryQtd.configure(highlightbackground="#d9d9d9")
        self.EntryQtd.configure(highlightcolor="#000000")
        self.EntryQtd.configure(insertbackground="#000000")
        self.EntryQtd.configure(selectbackground="#d9d9d9")
        self.EntryQtd.configure(selectforeground="black")
        self.EntryQtd.configure(validate='all')
        self.EntryQtd.configure(validatecommand=(vcmd, '%P'))
        self.EntryQtd.configure(textvariable=self.qtd)

        self.ButtonAtualizar = tk.Button(self.LabelframeQtd)
        self.ButtonAtualizar.place(relx=0.714, rely=0.128, height=36, width=97
                , bordermode='ignore')
        self.ButtonAtualizar.configure(activebackground="#d9d9d9")
        self.ButtonAtualizar.configure(activeforeground="black")
        self.ButtonAtualizar.configure(background="#d9d9d9")
        self.ButtonAtualizar.configure(disabledforeground="#a3a3a3")
        self.ButtonAtualizar.configure(foreground="#000000")
        self.ButtonAtualizar.configure(highlightbackground="#d9d9d9")
        self.ButtonAtualizar.configure(highlightcolor="#000000")
        self.ButtonAtualizar.configure(state='disabled')
        self.ButtonAtualizar.configure(text='''Atualizar''')
        self.ButtonAtualizar.configure(command=lambda:self.button_atualizar_callback())

        self.CanvasLotesQtd = ctk.CTkScrollableFrame(self.LabelframeQtd)
        self.CanvasLotesQtd.place(relx=0.048, rely=0.253, relheight=0.691
                , relwidth=0.9, bordermode='ignore')
        self.CanvasLotesQtd.configure(fg_color="#d9d9d9")
        self.CanvasLotesQtd.configure(corner_radius=0)
        self.CanvasLotesQtd.configure(border_width=1)
        #self.CanvasLotesQtd.configure(highlightbackground="#d9d9d9")
        #self.CanvasLotesQtd.configure(highlightcolor="#000000")
        #self.CanvasLotesQtd.configure(insertbackground="#000000")
        #self.CanvasLotesQtd.configure(relief="ridge")
        #self.CanvasLotesQtd.configure(selectbackground="#d9d9d9")
        #self.CanvasLotesQtd.configure(selectforeground="black")
        
        self.FramesLote = []
        self.LabelsLote = []
        self.EntriesLote = []

        #load_data()