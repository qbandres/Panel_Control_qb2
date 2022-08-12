from operator import index
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import numpy as np
import pandas as pd
from tkinter.filedialog import asksaveasfile
import datetime as dt
from datetime import date, timedelta

from package.function import PB_import_master, PB_import_otec, PB_export, pte_import_master, pte_import_pte, pte_ejec, pte_export, pte_export_tekla, sist_import_piping,sist_import_mec, sist_import_elect,sist_import_steel, sist_import_arq, sist_import_OC, sist_import_list, sist_export_data, sist_export_report_bi, asistencia, export_asis, import_id

class Widget:
    def __init__(self, fram, back, ancho, altura, pox, poy):
        self.fram = fram
        self.back = back
        self.pox = pox
        self.poy = poy
        self.altura = altura
        self.ancho = ancho

    def boton(self, name, action):
        Button(self.fram, text=name, bg=self.back, width=self.ancho, height=self.altura, command=action).place(
            x=self.pox, y=self.poy)

    def marco(self):
        Frame(self.fram, bg=self.back, width=self.ancho, height=self.altura, relief='sunken', bd=2).place(
            x=self.pox, y=self.poy)

    def letra(self, name):
        Label(self.fram, text=name, bg=self.back, padx=self.ancho, pady=self.altura).place(x=self.pox,y=self.poy)

root = Tk()
root.title('Control Panel')
root.geometry("380x420")

my_notebook = ttk.Notebook(root)
my_notebook.pack()

my_frame1 = Frame(my_notebook,width=500,height=500,bg="gray77")
my_frame2 = Frame(my_notebook,width=500,height=500,bg="gray77")
my_frame3 = Frame(my_notebook,width=500,height=500,bg="gray")
my_frame4 = Frame(my_notebook,width=500,height=500,bg="gray")
my_frame5 = Frame(my_notebook,width=500,height=500,bg="gray")

my_frame1.pack(fill ="both", expand=1)
my_frame2.pack(fill ="both", expand=1)
my_frame3.pack(fill ="both", expand=1)
my_frame4.pack(fill ="both", expand=1)
my_frame5.pack(fill ="both", expand=1)

my_notebook.add(my_frame1,text="PB Steel")
my_notebook.add(my_frame2,text="Avan Pte")
my_notebook.add(my_frame3,text="System")
my_notebook.add(my_frame4,text="Asitencia")
my_notebook.add(my_frame5,text="Tekla_paint")

#Frame 1
Widget(my_frame1,"gray56", 15, 2, 250, 5).boton('Importar Master',PB_import_master)
Widget(my_frame1,"gray56", 15, 2, 250, 45).boton('Importar OTEC',PB_import_otec)
Widget(my_frame1,"gray56", 15, 2, 250, 85).boton('Exportar PBI',PB_export)

#Frame 2
Widget(my_frame2,"gray56", 15, 1, 250, 5).boton('Importar Master',pte_import_master)
Widget(my_frame2,"gray77", 1, 1, 150, 5).letra('Importado')
Widget(my_frame2,"gray56", 15, 1, 250, 45).boton('Importar filtro',pte_import_pte)
Widget(my_frame2,"gray77", 1, 1, 150, 45).letra('Importado')

Widget(my_frame2,"gray56", 15, 1, 250, 85).boton('Procesar',pte_ejec)
Widget(my_frame2,"gray77", 1, 1, 150, 85).letra('Procesado')
Widget(my_frame2,"gray56", 15, 1, 250, 125).boton('Exportar Report pte',pte_export)
Widget(my_frame2,"gray77", 1, 1, 150, 125).letra('Exportado')
Widget(my_frame2,"gray56", 15, 1, 250, 165).boton('Exportar avance',pte_export_tekla)
Widget(my_frame2,"gray77", 1, 1, 150, 165).letra('Exportado')

#Frame3
Widget(my_frame3,"gray56", 15, 1, 250, 5).boton('Importar Piping',sist_import_piping)
Widget(my_frame3,"gray77", 15, 1, 150, 5).letra('Importado')
Widget(my_frame3,"gray56", 15, 1, 250, 40).boton('Importar Mec',sist_import_mec)
Widget(my_frame3,"gray77", 15, 1, 150, 40).letra('Importado')
Widget(my_frame3,"gray56", 15, 1, 250, 75).boton('Importar Elect',sist_import_elect)
Widget(my_frame3,"gray77", 15, 1, 150, 75).letra('Importado')
Widget(my_frame3,"gray56", 15, 1, 250, 110).boton('Import Steel',sist_import_steel)
Widget(my_frame3,"gray77", 15, 1, 150, 110).letra('Importado')
Widget(my_frame3,"gray56", 15, 1, 250, 145).boton('Import ARQ',sist_import_arq)
Widget(my_frame3,"gray77", 15, 1, 150, 145).letra('Importado')
Widget(my_frame3,"gray56", 15, 1, 250, 180).boton('Import OOCC',sist_import_OC)
Widget(my_frame3,"gray77", 15, 1, 150, 180).letra('Importado')

Widget(my_frame3,"gray56", 15, 1, 250, 215).boton('Import Listado Sist',sist_import_list)
Widget(my_frame3,"gray77", 15, 1, 150, 215).letra('Importado')
Widget(my_frame3,"gray56", 15, 1, 250, 250).boton('Exportar Data',sist_export_data)
Widget(my_frame3,"gray77", 1, 1, 150, 250).letra('OK')
Widget(my_frame3,"gray56", 15, 1, 250, 285).boton('Exportar PBI',sist_export_report_bi)
Widget(my_frame3,"gray77", 1, 1, 150, 285).letra('OK')

#Frame4
Widget(my_frame4,"gray56",15,1,250,5).boton("Importar asistencia",asistencia)
Widget(my_frame4,"gray77", 1, 1, 150, 5).letra('OK')
Widget(my_frame4,"gray56",15,1,250,40).boton("Exportar asistencia",export_asis)
Widget(my_frame4,"gray77", 1, 1, 150, 40).letra('OK')

#Frame5
Widget(my_frame5,"gray56",15,1,250,5).boton("Importar listado ID",import_id)
Widget(my_frame5,"gray77", 1, 1, 150, 40).letra('OK')

# Widget(my_frame3,"gray56", 15, 1, 250, 45).boton('Export PBI',sist_export_pbi)

root.mainloop()