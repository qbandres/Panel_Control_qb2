import sys
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import numpy as np
import pandas as pd
from tkinter.filedialog import asksaveasfile

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
class Semana:                                   #CREAR DATA FRAME CON LA SEMANA
    def __init__(self,df):
        self.fi='2019-04-12'
        self.T=1300
        self.df=df

    def split(self):
        s = pd.date_range(start=self.fi, periods=self.T, freq='D') # Creas el ranfo de fechas
        Nsemana = pd.DataFrame(s, columns=['Fecha'])  # Lo conveiertes en dataframe
        Nsemana['SEMANA'] = Nsemana.index
        Nsemana["Fecha"] = pd.to_datetime(Nsemana.Fecha).dt.date
        Nsemana.set_index('Fecha', inplace=True)
        Nsemana['Semana'] = Nsemana.SEMANA // 7 + 1
        Nsemana['NSem'] = Nsemana.SEMANA % 7 + 1
        del Nsemana['SEMANA']
        Nsemana['FECHA'] = Nsemana.index
        Nsemana.reset_index(drop=True, inplace=True)
        Nsemana = Nsemana[['FECHA', 'Semana','NSem']]

        self.df = self.df.merge(Nsemana, on='FECHA', how='left')  # Buscas HH de ID y lo insertas en data Horas ganadas
        return self.df

#Funiones de Power bi - Steel
def PB_import():
    global dfv, df_base

    ########CODIGO STEEL###############

    import_file_path = filedialog.askopenfilename()
    df_master = pd.read_excel(import_file_path,sheet_name='Reporte',skiprows=7)
    # df_otec = pd.read_excel(import_file_path,sheet_name='otec')

    # print(df_otec)

    d_pon = {'TR': 0.05, 'PA': 0.1, 'MO': 0.45, 'NI': 0.2, 'PI': 0.1, 'PU': 0.1}  # PONDERACIONES STEEL
    df_master = df_master[['IDTekla', 'ESP', 'Barcode', 'PesoTotal(Kg)', 'Ratio', 'Traslado', 'Prearmado', 'Montaje',
                           'Nivelacion,soldadura&Torque', 'Punchlist', 'FASE', 'Clasificación','linea']]

    df_master.rename(columns={'Traslado': 'DTR', 'Prearmado': 'DPA',
                              'Montaje': 'DMO', 'Nivelacion,soldadura&Torque': 'DNI',
                              'Punchlist': 'DPU', 'IDTekla': 'ID', 'PesoTotal(Kg)': 'WEIGHT'},
                     inplace=True)


    df_master = df_master[df_master.Ratio.notnull()]  # LIMPIAMOS DATOS QUE ESTEN NULOS EN EL RATIO
    df_master = df_master[df_master.WEIGHT.notnull()]  # LIMPIAMOS LOS DATOS QUE ESTEN NULOS EN EL PESO

    df_master = df_master[df_master["FASE"] != "PEBBLES"]
    # df = df[df["FASE"] != "PIPE RACK 1@2"]
    # df = df[df["FASE"] != "SOPORTES COMITIVA PIPERACK 3 CON SALAS ELECTRICAS"]
    df_master = df_master[df_master["FASE"] != "Torre de transferencia"]

    


    # # CALCULO DE PESOS TOTALES DE SEGUN PONDERACION

    df_master['TOTAL_WTR'] = df_master.WEIGHT * d_pon['TR']
    df_master['TOTAL_WPA'] = df_master.WEIGHT * d_pon['PA']
    df_master['TOTAL_WMO'] = df_master.WEIGHT * d_pon['MO']
    df_master['TOTAL_WNI'] = df_master.WEIGHT * d_pon['NI']
    df_master['TOTAL_WPU'] = df_master.WEIGHT * d_pon['PU']


    # # CALCULO DE HH EARNED TOTALES SEGUN MODERATION

    df_master['TOTAL_ETR'] = df_master.WEIGHT * d_pon['TR'] * df_master.Ratio / 1000
    df_master['TOTAL_EPA'] = df_master.WEIGHT * d_pon['PA'] * df_master.Ratio / 1000
    df_master['TOTAL_EMO'] = df_master.WEIGHT * d_pon['MO'] * df_master.Ratio / 1000
    df_master['TOTAL_ENI'] = df_master.WEIGHT * d_pon['NI'] * df_master.Ratio / 1000
    df_master['TOTAL_EPU'] = df_master.WEIGHT * d_pon['PU'] * df_master.Ratio / 1000

    


    # CALCULO DE PESO SEGUN AVANCE
    df_master['WTR'] = np.where(df_master['DTR'].isnull(), 0, df_master.WEIGHT * d_pon['TR'])
    df_master['WPA'] = np.where(df_master['DPA'].isnull(), 0, df_master.WEIGHT * d_pon['PA'])
    df_master['WMO'] = np.where(df_master['DMO'].isnull(), 0, df_master.WEIGHT * d_pon['MO'])
    df_master['WNI'] = np.where(df_master['DNI'].isnull(), 0, df_master.WEIGHT * d_pon['NI'])
    df_master['WPU'] = np.where(df_master['DPU'].isnull(), 0, df_master.WEIGHT * d_pon['PU'])

    # CALCULO DE PESO BRUTO QUIEBRE AVANCE
    df_master['BWTR'] = np.where(df_master['DTR'].isnull(), 0, df_master.WEIGHT)
    df_master['BWPA'] = np.where(df_master['DPA'].isnull(), 0, df_master.WEIGHT)
    df_master['BWMO'] = np.where(df_master['DMO'].isnull(), 0, df_master.WEIGHT)
    df_master['BWNI'] = np.where(df_master['DNI'].isnull(), 0, df_master.WEIGHT)
    df_master['BWPU'] = np.where(df_master['DPU'].isnull(), 0, df_master.WEIGHT)



    df_base = df_master[['ID', 'ESP', 'WEIGHT', 'Ratio', 'FASE', 'Clasificación','linea', 'BWTR', 'BWPA', 'BWMO', 'BWNI', 'BWPU']]

    df_base = df_base.fillna(0)


    df_base['WEIGHT'] = df_base['WEIGHT']* 0.001
    df_base.loc['BWPA'] = df_base['BWPA'] * 0.001
    df_base.loc['BWMO'] = df_base['BWMO'] * 0.001
    df_base.loc['BWNI'] = df_base['BWNI'] * 0.001
    df_base.loc['BWPU'] = df_base['BWPU'] * 0.001



    df_base.dropna(inplace=True)
    df_base.rename(columns={'FASE': 'ZONA'},
                       inplace=True)


    # CALCULO DE HH EARNED SEGUN AVANCE
    df_master['ETR'] = np.where(df_master['DTR'].isnull(), 0, df_master.WEIGHT * d_pon['TR'] * df_master.Ratio / 1000)
    df_master['EPA'] = np.where(df_master['DPA'].isnull(), 0, df_master.WEIGHT * d_pon['PA'] * df_master.Ratio / 1000)
    df_master['EMO'] = np.where(df_master['DMO'].isnull(), 0, df_master.WEIGHT * d_pon['MO'] * df_master.Ratio / 1000)
    df_master['ENI'] = np.where(df_master['DNI'].isnull(), 0, df_master.WEIGHT * d_pon['NI'] * df_master.Ratio / 1000)
    df_master['EPU'] = np.where(df_master['DPU'].isnull(), 0, df_master.WEIGHT * d_pon['PU'] * df_master.Ratio / 1000)

    df_master['WBRUTO'] = np.where(df_master['DMO'].isnull(), 0, df_master.WEIGHT)
    df_master['WPOND'] = df_master.WTR + df_master.WPA + df_master.WMO + df_master.WNI + df_master.WPU




    #########################SEPARAMOS LOS PESOS POR AVANCE DE CADA ETAPA

    df_dtr = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DTR", "WTR", "ETR", 'FASE', 'Clasificación','linea']]
    df_dtr = df_dtr.dropna(subset=['DTR'])  # Elimina llas filas vacias de DTR
    df_dtr["Etapa"] = "1-Traslado"
    df_dtr = df_dtr.rename(columns={'WTR': 'WPOND', "DTR": 'Fecha', 'ETR': 'HGan'})

    df_dpa = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DPA", "WPA", "EPA", 'FASE', 'Clasificación','linea']]
    df_dpa = df_dpa.dropna(subset=['DPA'])
    df_dpa["Etapa"] = "2-Ensamble"
    df_dpa = df_dpa.rename(columns={'WPA': 'WPOND', "DPA": 'Fecha', 'EPA': 'HGan'})

    df_dmo = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DMO", "WMO", "EMO", 'FASE', 'Clasificación','linea']]
    df_dmo = df_dmo.dropna(subset=['DMO'])
    df_dmo["Etapa"] = "3-Montaje"
    df_dmo = df_dmo.rename(columns={'WMO': 'WPOND', "DMO": 'Fecha', 'EMO': 'HGan'})

    df_dni = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DNI", "WNI", "ENI", 'FASE', 'Clasificación','linea']]
    df_dni = df_dni.dropna(subset=['DNI'])
    df_dni["Etapa"] = "4-Alineamiento"
    df_dni = df_dni.rename(columns={'WNI': 'WPOND', "DNI": 'Fecha', 'ENI': 'HGan'})

    df_dpu = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DPU", "WPU", "EPU", 'FASE', 'Clasificación','linea']]
    df_dpu = df_dpu.dropna(subset=['DPU'])
    df_dpu["Etapa"] = "6-Punch_List"
    df_dpu = df_dpu.rename(columns={'WPU': 'WPOND', "DPU": 'Fecha', 'EPU': 'HGan'})

    #CONCATENAR VERTICAL DE LAS COLUMNAS DE RESUMEN

    dfv = pd.concat(
        [df_dtr.round(1), df_dpa.round(1), df_dmo.round(1), df_dni.round(1), df_dpu.round(1)], axis=0)

    dfv.to_excel('mmm.xlsx')


    dfv['WBRUTO'] = np.where(dfv.Etapa != '3-Montaje', 0, dfv.WEIGHT)

    np_array = dfv.to_numpy()
    dfv = pd.DataFrame(data=np_array,
                       columns=['ESP', 'ID', 'Barcode', 'WEIGHT', 'Ratio', 'FECHA', 'WPOND', 'HHGan', 'FASE',
                                'Clasificación','linea', 'Etapa',
                                'WBRUTO'])
    dfv["FECHA"] = pd.to_datetime(dfv.FECHA).dt.date

    

    dfv = Semana(dfv).split()  # Insertamos la Semana con class
    print(dfv)
    dfv.to_excel('mmm.xlsx')

    dfv = dfv.fillna(0)

    print(df_master)
    print(df_base)
    print(dfv.WPOND)

    dfv['WPOND'] = dfv['WPOND']*0.001
    


    print(dfv)
 


    dfv['WBRUTO'] = dfv['WBRUTO']*0.001
    dfv['WEIGHT'] = dfv['WEIGHT'] * 0.001

    dfv.dropna(subset=['HHGan'], inplace=True)
    dfv['Disc'] = 'Steel'

    dfv.rename(columns={'FASE': 'ZONA'},
                       inplace=True)

    Widget(root,"gray77", 1, 1, 140, 38).letra('STEEL-M')
def PB_export():

    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'QB2_STEEL.xlsx')  # Creamos una excel y le indicamos la ruta

    # Exportar Steel
    dfv.to_excel(writer, sheet_name='ST_Gan', index=True)
    df_base.to_excel(writer, sheet_name='ST_Base', index=True)

    Widget(root,"gray77", 1, 1, 168, 178).letra('OK')

    writer.save()


#Funcion Avanve Steel
def pte_import_master():
    global mast
    #LECTURA DE LA INFORMACION

    import_file_path = filedialog.askopenfilename()
    mast = pd.read_excel(import_file_path,sheet_name='Reporte',skiprows=7)

    #RENOMBRANDO LAS COLUMNAS
    mast.rename(columns={"PesoUnit,(Kg)":"Weight","Traslado":"DTR","Prearmado":"DEN","Montaje":"DMO","Nivelacion,soldadura&Torque":"DTO","Punchlist":"DPU","IDTekla":"ID"},inplace=True)

    #RETIRARNO LAS FASES QUE NO INTERVIENEN
    mast = mast[mast["FASE"] != "PEBBLES"]
    mast = mast[mast["FASE"] != "PIPE RACK 1@2"]
    mast = mast[mast["FASE"] != "SOPORTES COMITIVA PIPERACK 3 CON SALAS ELECTRICAS"]
    mast = mast[mast["FASE"] != "Torre de transferencia"]

    #FILTRANDO LAS COLUMNAS QUE SE VAN A USAR
    mast = mast[['ESP','Clasificación','Stockcode','Descripción','Barcode','ID','Weight','DTR','DEN','DMO','DTO','DPU']]      

    Widget(my_frame2,"gray77", 1, 1, 150, 5).letra('Importado')
def pte_import_pte():
    global pte_1

    import_file_path = filedialog.askopenfilename()
    pte_1 = pd.read_excel(import_file_path,sheet_name='data')

    Widget(my_frame2,"gray77", 1, 1, 150, 45).letra('Importado')
def pte_ejec():

    global pte_res, pte_res_f, pte_res_ft,pte , avan

    avan = mast.copy()
    pte = pte_1.copy()

    avan['USER_FIELD_1'] = np.where(avan['DMO'].isnull(),'Falta','Montaje')
    avan['USER_FIELD_1'] = np.where(avan['DTO'].isnull(),avan.USER_FIELD_1,'Torque')
    # avan.to_excel('Avance total.xlsx')

    #SE FILTRA LO QUE FALTA
    #avan = avan[avan["USER_FIELD_1"] != "Falta"]

    #CREANDO LA LISTA DE ACTUALIZACIÓN DE LOS PUENTES GRÚA
    pte = pte.merge(mast[['ID','DMO','DTO']], on='ID',how='left')
    pte['Est_Mont'] = np.where(pte['DMO'].isnull(), 0, pte.WEIGHT)
    pte['Est_Torqu'] = np.where(pte['DTO'].isnull(), 0, pte.WEIGHT)

    pte['Pza_Saldo_Montaje'] = np.where(pte['DMO'].isnull(),1,0)
    pte['Pza_saldo_Torque'] = np.where(pte['DTO'].isnull(),1,0)
    pte['USER_FIELD_1'] = np.where(pte['DMO'].isnull(),'Falta','Montaje')
    pte['USER_FIELD_1'] = np.where(pte['DTO'].isnull(),pte.USER_FIELD_1,'Torque')

    #CREANDO FILTRO SOBRE CONDICION 1
    pte_f = pte[pte['CONDICION1']=='Funcionamiento de puente grua']

    #AGRUPAMOS PARA LOS AVANCES 1
    pte_res = pte.groupby(['SECTOR_MONTAJE']).sum()
    pte_res = pte_res[['WEIGHT','Est_Mont','Est_Torqu','QUANTITY','Pza_Saldo_Montaje','Pza_saldo_Torque']]
    pte_res['A_Mont'] = pte_res.Est_Mont/pte_res.WEIGHT
    pte_res['A_Torqu'] = pte_res.Est_Torqu/pte_res.WEIGHT
    pte_res.reset_index(level=0, inplace=True)

    #AGRUPAMOS PARA LOS AVANCES 2
    pte_res_f = pte_f.groupby(['SECTOR_MONTAJE']).sum()
    pte_res_f = pte_res_f[['WEIGHT','Est_Mont','QUANTITY','Est_Torqu','Pza_Saldo_Montaje','Pza_saldo_Torque']]
    pte_res_f['Saldo_Mont'] = pte_res_f.WEIGHT - pte_res_f.Est_Mont
    pte_res_f['A_Mont'] = pte_res_f.Est_Mont/pte_res_f.WEIGHT
    pte_res_f['A_Torqu'] = pte_res_f.Est_Torqu/pte_res_f.WEIGHT
    pte_res_f.reset_index(level=0, inplace=True)

    #AGRUPAMOS PARA LE FUNCIONAMIENTO DEL TORQUE
    pte_res_ft = pte_f.groupby(['SECTOR_TORQUE']).sum()
    pte_res_ft = pte_res_ft[['WEIGHT','Est_Mont','Est_Torqu','QUANTITY','Pza_Saldo_Montaje','Pza_saldo_Torque']]
    pte_res_ft['Saldo_Torque'] = pte_res_ft.WEIGHT - pte_res_ft.Est_Torqu
    pte_res_ft['A_Mont'] = pte_res_ft.Est_Mont/pte_res_ft.WEIGHT
    pte_res_ft['A_Torqu'] = pte_res_ft.Est_Torqu/pte_res_ft.WEIGHT
    pte_res_ft.reset_index(level=0, inplace=True)

    Widget(my_frame2,"gray77", 1, 1, 150, 85).letra('Procesado')
def pte_export():
    
    #Exportar
    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'Exp_Avance_pte.xlsx')  # Creamos una excel y le indicamos la ruta

    pte.to_excel(writer, sheet_name='Detalle', index=False)
    pte_res.to_excel(writer, sheet_name='Resumen_total', index=False)
    pte_res_f.to_excel(writer, sheet_name='Resumen_Func_mont', index=False)
    pte_res_ft.to_excel(writer, sheet_name='Resumen_Func_torque', index=False)

    writer.save()

    Widget(my_frame2,"gray77", 1, 1, 150, 125).letra('Exportado')
def pte_export_tekla():

    avan_1 = avan.copy()

    avan_1 = avan_1[['ID','USER_FIELD_1']]
    avan_1.columns=['ID','USER_FIELD_3']
    avan_1 = avan_1.dropna()


    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    avan_1.to_csv(export_file + '/' + 'pintado_avance.csv',index=False)  # Creamos una excel y le indicamos la ruta

    Widget(my_frame2,"gray77", 1, 1, 150, 165).letra('Exportado')

#Función Systemas
def sys_import_piping():
    import_file_path = filedialog.askopenfilename()
    pip_sys = pd.read_excel(import_file_path,sheet_name='DATA')
    print(pip_sys)
    print(pip_sys.columns)

    Widget(my_frame2,"gray77", 15, 1, 250, 5).letra('Importado')

root = Tk()
root.title('Control Panel')
root.geometry("380x235")

my_notebook = ttk.Notebook(root)
my_notebook.pack()

my_frame1 = Frame(my_notebook,width=500,height=500,bg="gray77")
my_frame2 = Frame(my_notebook,width=500,height=500,bg="gray77")
my_frame3 = Frame(my_notebook,width=500,height=500,bg="gray")

my_frame1.pack(fill = "both", expand=1)
my_frame2.pack(fill="both", expand=1)
my_frame3.pack(fill="both", expand=1)

my_notebook.add(my_frame1,text="PB Steel")
my_notebook.add(my_frame2,text="Avan Pte")
my_notebook.add(my_frame3,text="System")

#Frame 1
Widget(my_frame1,"gray56", 15, 6, 250, 5).boton('Importar Master',PB_import)
Widget(my_frame1,"gray56", 15, 6, 250, 105).boton('Exportar PBI',PB_export)


#Frame 2
Widget(my_frame2,"gray56", 15, 1, 250, 5).boton('Importar Master',pte_import_master)
Widget(my_frame2,"gray56", 15, 1, 250, 45).boton('Importar Report_pte',pte_import_pte)
Widget(my_frame2,"gray56", 15, 1, 250, 85).boton('Procesar',pte_ejec)
Widget(my_frame2,"gray56", 15, 1, 250, 125).boton('Exportar Report pte',pte_export)
Widget(my_frame2,"gray56", 15, 1, 250, 165).boton('Exportar avance',pte_export_tekla)


#Frame3
Widget(my_frame3,"gray56", 15, 1, 250, 5).boton('Importar Piping',sys_import_piping)

root.mainloop()