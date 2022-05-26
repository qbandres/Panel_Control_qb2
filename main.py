from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import numpy as np
import pandas as pd
from tkinter.filedialog import asksaveasfile
import datetime as dt 
from datetime import date, timedelta

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
        self.T=2500
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
def PB_import_master():
    global dfv, df_base,dfv_g

    ########CODIGO STEEL###############

    import_file_path = filedialog.askopenfilename()
    df_master = pd.read_excel(import_file_path,sheet_name='Reporte',skiprows=7)
    # df_otec = pd.read_excel(import_file_path,sheet_name='otec')

    # print(df_otec)

    d_pon = {'TR': 0.05, 'PA': 0.1, 'MO': 0.45, 'NI': 0.2, 'PI': 0.1, 'PU': 0.1}  # PONDERACIONES STEEL
    df_master = df_master[['IDTekla', 'ESP', 'Barcode', 'PesoTotal(Kg)', 'Ratio', 'Traslado', 'Prearmado', 'Montaje',
                           'Nivelacion,soldadura&Torque', 'Punchlist', 'FASE', 'Clasificación','linea','Descripción']]

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

    df_master["DTR"] = pd.to_datetime(df_master.DTR).dt.date
    df_master["DPA"] = pd.to_datetime(df_master.DPA).dt.date
    df_master["DMO"] = pd.to_datetime(df_master.DMO).dt.date
    df_master["DNI"] = pd.to_datetime(df_master.DNI).dt.date
    df_master["DPU"] = pd.to_datetime(df_master.DPU).dt.date

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

    df_base = df_master[['ID', 'ESP', 'WEIGHT', 'Ratio', 'FASE', 'Clasificación','linea','Descripción', 'BWTR', 'BWPA', 'BWMO', 'BWNI', 'BWPU']]

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

    df_dtr = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DTR", "WTR", "ETR", 'FASE', 'Clasificación','linea','Descripción']]
    df_dtr = df_dtr.dropna(subset=['DTR'])  # Elimina llas filas vacias de DTR
    df_dtr["Etapa"] = "1-Traslado"
    df_dtr = df_dtr.rename(columns={'WTR': 'WPOND', "DTR": 'Fecha', 'ETR': 'HGan'})

    df_dpa = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DPA", "WPA", "EPA", 'FASE', 'Clasificación','linea','Descripción']]
    df_dpa = df_dpa.dropna(subset=['DPA'])
    df_dpa["Etapa"] = "2-Ensamble"
    df_dpa = df_dpa.rename(columns={'WPA': 'WPOND', "DPA": 'Fecha', 'EPA': 'HGan'})

    df_dmo = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DMO", "WMO", "EMO", 'FASE', 'Clasificación','linea','Descripción']]
    df_dmo = df_dmo.dropna(subset=['DMO'])
    df_dmo["Etapa"] = "3-Montaje"
    df_dmo = df_dmo.rename(columns={'WMO': 'WPOND', "DMO": 'Fecha', 'EMO': 'HGan'})

    df_dni = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DNI", "WNI", "ENI", 'FASE', 'Clasificación','linea','Descripción']]
    df_dni = df_dni.dropna(subset=['DNI'])
    df_dni["Etapa"] = "4-Alineamiento"
    df_dni = df_dni.rename(columns={'WNI': 'WPOND', "DNI": 'Fecha', 'ENI': 'HGan'})

    df_dpu = df_master[["ESP", "ID", 'Barcode', "WEIGHT", "Ratio", "DPU", "WPU", "EPU", 'FASE', 'Clasificación','linea','Descripción']]
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
                                'Clasificación','linea','Descripción', 'Etapa',
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

    dfv_g = dfv.groupby(['ESP']).sum()

    Widget(root,"gray77", 1, 1, 140, 38).letra('STEEL-M')
def PB_import_otec():
    global df_otec,dg_otec

    import_file_path = filedialog.askopenfilename()
    df_otec= pd.read_excel(import_file_path,sheet_name='Estructuras',skiprows=6)

    df_otec = df_otec.iloc[:38505]    #Filtramos solo las columnas necesarias

    df_otec=df_otec[['ESP','Descripción','Peso Total (Kg)','Traslado.1','Pre-Armado.1','Montaje.1', 'Nivelación, Soldadura & Torque.1',
            'Punch List','total hh','traslado hh','prearm hh','montaje hh','niv, sold torq hh','punch list hh']]

    df_otec = df_otec[df_otec["ESP"] != "03201.H028"]
    df_otec = df_otec[df_otec["ESP"] != "03201.H029"]
    df_otec = df_otec[df_otec["ESP"] != "03201.H006"]
    df_otec = df_otec[df_otec["ESP"] != 1]
    df_otec['ESP'] =df_otec['ESP'].str.strip()


    dg_otec = df_otec.groupby('ESP').sum()

    df_otec = df_otec.fillna(0)

    print(df_otec)
    print(df_otec.columns)

    Widget(root,"gray77", 1, 1, 140, 118).letra('STEEL-OTEC')
def PB_export():

    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'QB2_STEEL.xlsx')  # Creamos una excel y le indicamos la ruta

    esp_comp = pd.concat([dg_otec,dfv_g],axis=1)



    # Exportar Steel
    dfv.to_excel(writer, sheet_name='ST_Gan', index=True)
    df_base.to_excel(writer, sheet_name='ST_Base', index=True)
    df_otec.to_excel(writer, sheet_name='ST_Otec', index=True)
    esp_comp.to_excel(writer, sheet_name='Comparacion_ESP', index=True)

    Widget(my_frame1,"gray77", 1, 1, 168, 178).letra('OK')

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
def sist_import_piping():


    global pip_tot, pip_tot_res,pip_conc

    import_file_path = filedialog.askopenfilename()
    pip_sist_1 = pd.read_excel(import_file_path,sheet_name='Cub General ',skiprows=10)
    pip_sist_2 = pd.read_excel(import_file_path,sheet_name='Cub Linea Menor',skiprows=9)
    pip_sist_3 = pd.read_excel(import_file_path,sheet_name='Cub Valvulas',skiprows=9)
    pip_sist_4 = pd.read_excel(import_file_path,sheet_name='Cub Soportes ',skiprows=4)

    print(pip_sist_1.columns)
    print(pip_sist_2.columns)
    print(pip_sist_3.columns)
    print(pip_sist_4.columns)

    #LECTURA DE PLANTILLAS OTEC

    pip_sist_gen = pip_sist_1[['COD','SUB-SISTEMA','Codigo Fluido','UG/AG/OR','ISOMETRICO','Metros Aislados','Diametro','Metros cañerias fore cast 11',
                        'AVANCE FINAL','HH TOTAL','HH\nAVANCE']]

    pip_sist_men = pip_sist_2[['SUB \nSISTEMA','Codigo Fluido','ISOMETRICOS','Metros Aislados','Diametro','Metros cañerias',
                        'AVANCE TOTAL\nAG','HH TOTAL','HH\nAVANCE']]
    
    pip_sist_valv = pip_sist_3[['SUB \nSISTEMA','Codigo Fluido','Isometrico','Diametro','Cantidad',
                         'AVANCE TOTAL\nAG','HH TOTAL','HH\nAVANCE']]
    pip_sist_soport = pip_sist_4[['Isometrico','Diametro','Peso soporte total KG','KG']]


    #CUBICACION GENERAL
    pip_sist_gen.rename(columns={'COD': 'QUIEBRE_OT','SUB-SISTEMA': 'SUBSISTEMA','HH TOTAL': 'HH_Tot','Metros cañerias fore cast 11': 'Cant_Tot','Diametro':'DIAMETRO','AVANCE FINAL':'Cant_Avan',
                                'HH\nAVANCE':'HH_Avan','ISOMETRICO':'TAG','UG/AG/OR':'Niv','Codigo Fluido':'Fluid'},
                       inplace=True)

    #INFO TEMPORAL PARA EXTRAR EL SUBSISTEMA
    iso_sub = pip_sist_gen[['TAG','SUBSISTEMA','Fluid']]
    iso_sub['ISO'] = iso_sub['TAG'].str[:13]
    del iso_sub['TAG']

    iso_sub = iso_sub.groupby(['ISO']).first()
    iso_sub = iso_sub.reset_index()

    print(iso_sub)

    iso_sub.to_excel('iso_sub.xlsx')

    pip_sist_gen = pip_sist_gen[pip_sist_gen['QUIEBRE_OT'] != 'FA']

    pip_sist_gen = pip_sist_gen.fillna(0)   

    pip_sist_gen['HH_Saldo'] = pip_sist_gen['HH_Tot'].subtract(pip_sist_gen["HH_Avan"])
    pip_sist_gen.insert(9, 'Cant_Sal', pip_sist_gen['Cant_Tot'].subtract(pip_sist_gen["Cant_Avan"]))

    pip_sist_gen['OTEC'] = 'General'
    pip_sist_gen['Und'] = 'ml'
    pip_sist_gen['ISO'] = pip_sist_gen['TAG'].str[:13]

    pip_sist_gen = pip_sist_gen[pip_sist_gen['QUIEBRE_OT'] != 0]



    #LINEA MENOR
    pip_sist_men.rename(columns={'SUB \nSISTEMA': 'SUBSISTEMA','HH TOTAL': 'HH_Tot','Metros cañerias': 'Cant_Tot','Diametro':'DIAMETRO','AVANCE TOTAL\nAG':'Cant_Avan',
                                'HH\nAVANCE':'HH_Avan','ISOMETRICOS':'TAG','Codigo Fluido':'Fluid'},inplace=True)

    pip_sist_men = pip_sist_men.fillna(0)

    pip_sist_men['HH_Saldo'] = pip_sist_men['HH_Tot'].subtract(pip_sist_men["HH_Avan"])
    
    pip_sist_men.insert(0, 'QUIEBRE_OT', 'Cañeria-SB (m)')
    pip_sist_men.insert(3, 'Niv', 'AG')

    pip_sist_men.insert(9, 'Cant_Sal', pip_sist_men['Cant_Tot'].subtract(pip_sist_men["Cant_Avan"]))

    pip_sist_men['OTEC'] = 'Linea_menor'
    pip_sist_men['Und'] = 'ml'
    pip_sist_men['ISO'] = pip_sist_men['TAG'].str[:13]

    pip_sist_men = pip_sist_men[pip_sist_men['QUIEBRE_OT'] != 0]


    #VAVULAS

    pip_sist_valv.rename(columns={'SUB \nSISTEMA': 'SUBSISTEMA','HH TOTAL': 'HH_Tot','Cantidad': 'Cant_Tot','Diametro':'DIAMETRO','AVANCE TOTAL\nAG':'Cant_Avan',
                                'HH\nAVANCE':'HH_Avan','Isometrico':'TAG','Codigo Fluido':'Fluid'},inplace=True)

    pip_sist_valv = pip_sist_valv.fillna(0)

    pip_sist_valv['HH_Saldo'] = pip_sist_valv['HH_Tot'].subtract(pip_sist_valv["HH_Avan"])
    
    pip_sist_valv.insert(0, 'QUIEBRE_OT', 'Cañería-LB (válvula)')
    pip_sist_valv.insert(3, 'Niv', 'AG')

    pip_sist_valv.insert(8, 'Cant_Sal', pip_sist_valv['Cant_Tot'].subtract(pip_sist_valv["Cant_Avan"]))

    pip_sist_valv['OTEC'] = 'Valvulas'
    pip_sist_valv['Und'] = 'Und'
    pip_sist_valv['ISO'] = pip_sist_valv['TAG'].str[:13]
    pip_sist_valv.insert(5, 'Metros Aislados', 'NA')


    #SOPORTERIA
    pip_sist_soport.rename(columns={'Peso soporte total KG': 'Cant_Tot','Diametro':'DIAMETRO','KG':'Cant_Avan','Isometrico':'TAG'},inplace=True)

    pip_sist_soport['ISO'] = pip_sist_soport['TAG'].str[:13]

    pip_sist_soport = pip_sist_soport.groupby(['ISO']).sum()

    pip_sist_soport = pip_sist_soport.reset_index()
    pip_sist_soport['HH_Tot'] = 0.42*pip_sist_soport.Cant_Tot
    pip_sist_soport['HH_Avan'] = 0.42*pip_sist_soport.Cant_Avan


    pip_sist_soport = pip_sist_soport.merge(iso_sub[['ISO','SUBSISTEMA','Fluid']], on='ISO', how='left')
    pip_sist_soport = pip_sist_soport.dropna(subset=['SUBSISTEMA'])


    pip_sist_soport.insert(0, 'QUIEBRE_OT', 'Cañería LB&SB (soporte)')
    pip_sist_soport = pip_sist_soport[['QUIEBRE_OT','SUBSISTEMA','Fluid','Cant_Tot','Cant_Avan','HH_Tot','HH_Avan','ISO']]

    pip_sist_soport.insert(3, 'Niv', 'AG')
    pip_sist_soport.insert(4, 'TAG', pip_sist_soport.ISO)
    pip_sist_soport.insert(5, 'DIAMETRO', 'NA')
    pip_sist_soport.insert(8, 'Cant_Sal', pip_sist_soport['Cant_Tot'].subtract(pip_sist_soport["Cant_Avan"]))
    pip_sist_soport.insert(11, 'HH_Saldo', pip_sist_soport['HH_Tot'].subtract(pip_sist_soport["HH_Avan"]))
    pip_sist_soport['OTEC'] = 'Valvulas'
    pip_sist_soport['Und'] = 'Und'
    del pip_sist_soport[ 'ISO']
    pip_sist_soport[ 'ISO'] = pip_sist_soport[ 'TAG']

    pip_sist_soport.insert(5, 'Metros Aislados', 'NA')

    print(pip_sist_soport)

    print(pip_sist_gen)
    print(pip_sist_men)
    print(pip_sist_valv.columns)

    #Aislamiento

    pip_aisl = pd.concat([pip_sist_gen,pip_sist_men],axis=0)
    pip_aisl['Cant_Tot'] = pip_aisl['Metros Aislados']
    pip_aisl['HH_Tot'] = pip_aisl['Cant_Tot']*4.54
    pip_aisl['HH_Avan'] = 0
    pip_aisl['Cant_Avan'] = 0
    pip_aisl['HH_Saldo'] = pip_aisl['HH_Tot']
    pip_aisl['Cant_Sal'] = pip_aisl['Cant_Tot']
    pip_aisl['Und'] = 'ml-Aisl'
    pip_aisl = pip_aisl.dropna(subset=['Cant_Tot'])

    pip_aisl['quiebre'] = np.where(pip_aisl.QUIEBRE_OT == 'Cañería-LB (m)','Cañería-LB (aislamiento)','Cañería-SB (aislamiento)')
    del pip_aisl['QUIEBRE_OT']
    pip_aisl.insert(0, 'QUIEBRE_OT',pip_aisl['quiebre'])
    del pip_aisl['quiebre']

    print(pip_aisl)

    #AGRUPAMIENTO FINAL
    pip_tot = pd.concat([pip_sist_gen,pip_sist_men,pip_sist_valv,pip_sist_soport,pip_aisl],axis=0)
    pip_tot['disc'] = 'Piping'
    pip_tot_res = pip_tot.groupby(['QUIEBRE_OT']).sum()

    pip_conc = pip_tot[['QUIEBRE_OT', 'SUBSISTEMA','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]


    Widget(my_frame3,"gray77", 15, 1, 150, 5).letra('Importado')
def sist_import_mec():
    global mec_sist, mec_conc
    import_file_path = filedialog.askopenfilename()
    mec_sist = pd.read_excel(import_file_path,sheet_name='Base Datos',skiprows=6)


    mec_sist = mec_sist[['TAG','SUBSIST','UND','HH UND','HH SALDO','Disciplina']]
    mec_sist.rename(columns={'SUBSIST': 'SUBSISTEMA','HH UND': 'HH_Tot','HH SALDO': 'HH_Saldo'},
                       inplace=True)

    mec_sist = mec_sist.dropna(subset=['UND'])


    def conditions(x):
        if x == 'EA':
            return "EQ (und)"
        elif x == 'EA#':
            return "EQ (und)"
        elif x == 'M':
            return "EQ (m)"
        elif x == 'MT':
            return "EQ (ton)"
        elif x == 'MT#':
            return "EQ (ton)"
        else:
            return "NA"

    func = np.vectorize(conditions)
    energy_class = func(mec_sist["UND"])

    mec_sist.insert(0, 'QUIEBRE_OT',energy_class)

    # mec_sist["QUIEBRE_OT"] = energy_class

    mec_sist = mec_sist.fillna(0)

    mec_sist.insert(4, 'HH_Avan' , mec_sist['HH_Tot'].subtract(mec_sist["HH_Saldo"]))
    mec_sist.insert(3, 'Cant_Tot' , mec_sist.HH_Tot)
    mec_sist.insert(4, 'Cant_Avan' , mec_sist.HH_Avan)
    mec_sist.insert(5, 'Cant_Sal' , mec_sist.HH_Saldo)

    mec_sist['OTEC'] = 'Base_Datos'
    mec_sist['disc'] = mec_sist.Disciplina
    del mec_sist['Disciplina']
    del mec_sist['UND']

    mec_conc = mec_sist

    del mec_conc['TAG']
    

    print(mec_sist)



    Widget(my_frame3,"gray77", 15, 1, 150, 40).letra('Importado')
def sist_import_elect():
    global elect_sist_cab,elec_con_cab,elec_con_al,elec_con_mall,elec_con_epc,elec_con_equ,elec_con_inst
    import_file_path = filedialog.askopenfilename()
    elect_sist_cab = pd.read_excel(import_file_path,sheet_name='Cables',skiprows=10)
    elect_sist_alu = pd.read_excel(import_file_path,sheet_name='Alumbrado',skiprows=10)
    elect_sist_malla = pd.read_excel(import_file_path,sheet_name='Malla a Tierra',skiprows=11)
    elect_sist_epc = pd.read_excel(import_file_path,sheet_name='EPC',skiprows=10)
    elect_sist_equ = pd.read_excel(import_file_path,sheet_name='Eq.salas.trafos',skiprows=15)
    elect_sist_inst = pd.read_excel(import_file_path,sheet_name='Instrumentos',skiprows=8)

    #CABLES

    elect_sist_cab=elect_sist_cab.iloc[:, lambda elect_sist_cab: [2,3,7,14,32,37,43,54,56,57,59]]
    elect_sist_cab.columns = ['SUBSISTEMA','Área','TAG','Cantidad','Tendido','Avance','Tipo cable','Linea','HH_TOTAL','HH_AVANCE','1ERCOBRE']

    elect_sist_cab.rename(columns={'Tipo cable': 'QUIEBRE_OT','HH_TOTAL': 'HH_Tot','Cantidad': 'Cant_Tot','Avance':'Cant_Avan',
                                'HH_AVANCE':'HH_Avan'},
                       inplace=True)

    elect_sist_cab = elect_sist_cab.dropna(subset=['QUIEBRE_OT'])
    elect_sist_cab = elect_sist_cab.dropna(how='all')


    elect_sist_cab['HH_Saldo'] = elect_sist_cab['HH_Tot'].subtract(elect_sist_cab["HH_Avan"])
    elect_sist_cab['Cant_Sal'] = elect_sist_cab['Cant_Tot'].subtract(elect_sist_cab["Cant_Avan"])
    elect_sist_cab['Und'] = 'ml'
    elect_sist_cab['OTEC'] = 'CABLES'
    elect_sist_cab['disc'] = 'ELECT'
   

    elec_con_cab = elect_sist_cab[['QUIEBRE_OT','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    
    print(elect_sist_cab['HH_Tot'].sum())
    print(elect_sist_cab['HH_Avan'].sum())
    print(elect_sist_cab['HH_Saldo'].sum())


    #Alumbrado

    elect_sist_alu=elect_sist_alu.iloc[:, lambda elect_sist_alu: [8,21,23,28,32,46,48,49]]
    elect_sist_alu.columns = ['TAG','Und','Cant_Tot','Tendido','Cant_Avan','QUIEBRE_OT','HH_Tot','HH_Avan']

    # elect_sist_alu['QUIEBRE_OT'] = elect_sist_alu['QUIEBRE_OT'].fillna(value='Por definir')
    elect_sist_alu = elect_sist_alu.dropna(subset=['QUIEBRE_OT'])
    elect_sist_alu = elect_sist_alu.dropna(how='all')

    elect_sist_alu['HH_Saldo'] = elect_sist_alu['HH_Tot'].subtract(elect_sist_alu["HH_Avan"])
    elect_sist_alu['Cant_Sal'] = elect_sist_alu['Cant_Tot'].subtract(elect_sist_alu["Cant_Avan"])
    elect_sist_alu['OTEC'] = 'ALUMBRADO'
    elect_sist_alu['disc'] = 'ELECT'
    elect_sist_alu['SUBSISTEMA'] = '0310-NZC-001'

    elec_con_al = elect_sist_alu[['QUIEBRE_OT','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]

    

    #MALLA

    elect_sist_malla=elect_sist_malla.iloc[:, lambda elect_sist_malla: [11,17,18,25,28,39,41,42]]
    elect_sist_malla.columns = ['TAG','Cant_Tot','Und','Tendido','Cant_Avan','QUIEBRE_OT','HH_Tot','HH_Avan']

    print(elect_sist_malla)

    elect_sist_malla = elect_sist_malla[elect_sist_malla['QUIEBRE_OT'] != 'Descope Aterramiento EPC']
    elect_sist_malla = elect_sist_malla[elect_sist_malla['QUIEBRE_OT'] != 'Descope Malla Tierra']
    elect_sist_malla = elect_sist_malla[elect_sist_malla['QUIEBRE_OT'] != 'Otro contrato Aterramiento EPC']

    print(elect_sist_malla)

    elect_sist_malla['HH_Saldo'] = elect_sist_malla['HH_Tot'].subtract(elect_sist_malla["HH_Avan"])
    elect_sist_malla['Cant_Sal'] = elect_sist_malla['Cant_Tot'].subtract(elect_sist_malla["Cant_Avan"])
    elect_sist_malla['OTEC'] = 'Malla a Tierra'
    elect_sist_malla['disc'] = 'ELECT'
    elect_sist_malla['SUBSISTEMA'] = '0310-NZC-001'

    elec_con_mall = elect_sist_malla[['QUIEBRE_OT','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]

    #EPC

    elect_sist_epc=elect_sist_epc.iloc[:, lambda elect_sist_epc: [9,16,17,24,26,42,44,45]]
    elect_sist_epc.columns = ['TAG','Cant_Tot','Und','Tendido','Cant_Avan','QUIEBRE_OT','HH_Tot','HH_Avan']

    elect_sist_epc = elect_sist_epc[elect_sist_epc['QUIEBRE_OT'] != 'Descope']
    elect_sist_epc = elect_sist_epc[elect_sist_epc['QUIEBRE_OT'] != 'Eliminado']


    elect_sist_epc['HH_Saldo'] = elect_sist_epc['HH_Tot'].subtract(elect_sist_epc["HH_Avan"])
    elect_sist_epc['Cant_Sal'] = elect_sist_epc['Cant_Tot'].subtract(elect_sist_epc["Cant_Avan"])
    elect_sist_epc['OTEC'] = 'EPC'
    elect_sist_epc['disc'] = 'ELECT'
    elect_sist_epc['SUBSISTEMA'] = '0310-NZC-001'

    elec_con_epc = elect_sist_epc[['QUIEBRE_OT','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    print(elec_con_epc)


    #EQUIPOS
    elect_sist_equ=elect_sist_equ.iloc[:, lambda elect_sist_equ: [4,8,15,16,23,27,43,44,46,48]]
    elect_sist_equ.columns = ['area','TAG','Und','Cant_Tot','Montaje','Cant_Avan','HH_Tot','HH_Avan','SUBSISTEMA','QUIEBRE_OT']

    elect_sist_equ = elect_sist_equ[elect_sist_equ['area'] != 320]
    elect_sist_equ['QUIEBRE_OT'] = elect_sist_equ['QUIEBRE_OT'].fillna('INDEFINIDO')

    elect_sist_equ['HH_Saldo'] = elect_sist_equ['HH_Tot'].subtract(elect_sist_equ["HH_Avan"])
    elect_sist_equ['Cant_Sal'] = elect_sist_equ['Cant_Tot'].subtract(elect_sist_equ["Cant_Avan"])
    elect_sist_equ['OTEC'] = 'Equipos electricos'
    elect_sist_equ['disc'] = 'ELECT'


    elec_con_equ = elect_sist_equ[['QUIEBRE_OT','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    print(elec_con_equ)

    Widget(my_frame3,"gray77", 15, 1, 150, 75).letra('Importado')

    #inSTRUMENTOS

    elect_sist_inst=elect_sist_inst.iloc[:, lambda elect_sist_inst: [11,3,7,30,37,39,48,49,55]] 
    elect_sist_inst.columns = ['area','TAG','SUBSISTEMA','Cant_Tot','Montaje','Cant_Avan','HH_Tot','HH_Avan','QUIEBRE_OT']
    elect_sist_inst = elect_sist_inst[elect_sist_inst['area'] != 320]
    elect_sist_inst['QUIEBRE_OT'] = elect_sist_inst['QUIEBRE_OT'].fillna('INDEFINIDO')
    elect_sist_inst['HH_Saldo'] = elect_sist_inst['HH_Tot'].subtract(elect_sist_inst["HH_Avan"])
    elect_sist_inst['Cant_Sal'] = elect_sist_inst['Cant_Tot'].subtract(elect_sist_inst["Cant_Avan"])
    elect_sist_inst['OTEC'] = 'Instrumentos'
    elect_sist_inst['disc'] = 'ELECT'
    elect_sist_inst['Und'] = 'UND'

    elec_con_inst = elect_sist_inst[['QUIEBRE_OT','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    print(elec_con_inst)

def sist_import_steel():
    global steel_conc,oocc_conc,arq_conc

    import_file_path = filedialog.askopenfilename()
    steel_sist= pd.read_excel(import_file_path,sheet_name='Steel')
    oocc_sist= pd.read_excel(import_file_path,sheet_name='OOCC')
    arq_sist= pd.read_excel(import_file_path,sheet_name='ARQ')
    steel_conc = steel_sist[['QUIEBRE_OT', 'SUBSISTEMA','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]

    oocc_conc = oocc_sist[['QUIEBRE_OT', 'SUBSISTEMA','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]

    arq_conc = arq_sist[['QUIEBRE_OT', 'SUBSISTEMA','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]




    Widget(my_frame3,"gray77", 15, 1, 150, 110).letra('Importado')


#Import Fechas
def sist_import_list():
    global list_sist, list_sist_1
    import_file_path = filedialog.askopenfilename()
    list_sist_1 = pd.read_excel(import_file_path,sheet_name='data')

    list_sist = list_sist_1[list_sist_1['ALCANCE'] == 'VP2']

    pd.to_datetime(list_sist.FIN).dt.date 
    pd.to_datetime(list_sist.Actual).dt.date 

    Widget(my_frame3,"gray77", 15, 1, 150, 110).letra('Importado')

#Power Bi
def sist_export_report_bi():

    
    sist_tot  = pd.concat([pip_sist_gen[['Disciplina','TAG','SUBSISTEMA','HH_TOTAL','HH_SALDO']],mec_sist[['Disciplina','TAG','SUBSISTEMA','HH_TOTAL','HH_SALDO']],
                        elect_sist_cab[['Disciplina','TAG','SUBSISTEMA','HH_TOTAL','HH_SALDO']]],axis=0)


    sist_tot_g = sist_tot.groupby(['SUBSISTEMA','Disciplina']).sum()
    sist_tot_g['Avance'] = 1-sist_tot_g.HH_SALDO/sist_tot_g.HH_TOTAL

    sist_tot_gt = sist_tot.groupby(['SUBSISTEMA']).sum()
    sist_tot_g =  sist_tot_g.reset_index(level=['SUBSISTEMA','Disciplina'])
    sist_tot_gt['Avance'] = 1-sist_tot_gt.HH_SALDO/sist_tot_gt.HH_TOTAL
    sist_tot_gt =  sist_tot_gt.reset_index(level=['SUBSISTEMA'])
    
    #Extraer los nombres de sistemas a los avances de disciplinas
    sist_tot_gs = sist_tot_g.merge(list_sist_1[['SUBSISTEMA', 'DESCRIP','LINEA','1ER_COBRE','Actual','FIN','Saldo_dias']], on='SUBSISTEMA',
                    how='left')

    #Extraer el listado de sistermas a los avances agrupado
    sist_tot_gts = sist_tot_gt.merge(list_sist_1[['SUBSISTEMA', 'DESCRIP','LINEA','1ER_COBRE','Actual','FIN','Saldo_dias']], on='SUBSISTEMA',
                    how='left')

    #Extraes los avances al listado de sistemas
    list_sist_g = list_sist.merge(sist_tot_g[['SUBSISTEMA','Disciplina','HH_TOTAL','HH_SALDO','Avance']], on='SUBSISTEMA',
                    how='left')
    list_sist_g['HH_Sem'] = list_sist_g.HH_SALDO/(list_sist_g['Saldo_dias'])*7

    n=0

    titulos_1 = ['ITEM', 'CODE', 'AREA', 'SUBSISTEMA', 'DESCRIP', 'CONTRATO', 'TIPO',
       'CONTRATO.1', 'DEF', 'LINEA', '1ER_COBRE', 'Actual', 'FIN',
       'Saldo_dias', 'Disciplina', 'HH_TOTAL', 'HH_SALDO', 'Avance', 'HH_Sem',
       'HH_Dia']

    df_temp = pd.DataFrame(columns=titulos_1)
 
    for i in list_sist_g.index:

        print(i)

        ft = [list_sist_g.iloc[n,11] + timedelta(days=d) for d in range((list_sist_g.iloc[n,12] - list_sist_g.iloc[n,11]).days + 1)]  # CREAMOS LA LISTA DE FECHAS
        dft = pd.DataFrame({'FECHA': ft})

        dft['HH_Dia'] = list_sist_g.iloc[n,16]/len(dft)
        dft['HH_Tot'] = list_sist_g.iloc[n,15]/len(dft)
        dftv = list_sist_g.iloc[[i]]
        dftv['FECHA'] = ft[0]
        dftv.set_index('FECHA',inplace =True )
        dft.set_index('FECHA',inplace =True)

        res = pd.concat([dftv,dft],axis=1)
        res = res.fillna(method='ffill')   
        df_temp = pd.concat([df_temp,res],axis=0)
        print(df_temp)
        print(n)
        n=n+1
    

    del df_temp['HH_TOTAL']
    del df_temp['HH_SALDO']
    del df_temp['Avance']
    del df_temp['HH_Sem']

    df_temp.reset_index(level=0, inplace=True)
    df_temp.rename(columns={"index": "Fecha"},inplace=True)
    df_temp['FECHA']=pd.to_datetime(df_temp.Fecha).dt.date   
    df_temp = Semana(df_temp).split()  

    df_temp['Disciplina'] = np.where(df_temp.Disciplina.isnull(),"NA",df_temp['Disciplina'])
    df_temp['1ER_COBRE'] =  np.where(df_temp['1ER_COBRE'].isnull(),"NA",df_temp['1ER_COBRE'])    
        

    #Extraes los avances al listado de sistemas
    list_sist_gt = list_sist.merge(sist_tot_gt[['SUBSISTEMA','HH_TOTAL','HH_SALDO','Avance']], on='SUBSISTEMA',
                    how='left')
    list_sist_gt['HH_Sem'] = list_sist_gt.HH_SALDO/(list_sist_gt['Saldo_dias'])*7
    sist_tot = sist_tot.merge(list_sist_1[['SUBSISTEMA', 'DESCRIP','LINEA','1ER_COBRE','Actual','FIN','Saldo_dias']], on='SUBSISTEMA',
                    how='left')


    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'Sistemas_QB2.xlsx')  # Creamos una excel y le indicamos la ruta

    # Exportar Steel
    sist_tot_gs.to_excel(writer, sheet_name='Det_Disc_sist', index=False)
    sist_tot_gts.to_excel(writer, sheet_name='Agru_Disc_Sist', index=True)
    list_sist_g.to_excel(writer, sheet_name='Det_Sist_Disc', index=False)
    list_sist_gt.to_excel(writer, sheet_name='Agru_Sist_Disc', index=False)
    df_temp.to_excel(writer, sheet_name='PowerBi', index=False)
    sist_tot.to_excel(writer, sheet_name='PowerBi_1', index=False)




    Widget(my_frame3,"gray77", 1, 1, 150, 145).letra('OK')

    writer.save()


#reporte Excel
def sist_export_report():

    export = pd.concat([pip_conc,mec_conc,elec_con_cab,elec_con_al,elec_con_mall,elec_con_epc,elec_con_equ,elec_con_inst,steel_conc,oocc_conc,arq_conc],axis=0)
    export_agr = export.groupby(['disc','OTEC','QUIEBRE_OT']).sum()

    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'Saldos de obra SS.xlsx')  # Creamos una excel y le indicamos la 
    
    # Exportar Steel
    export.to_excel(writer, sheet_name='Det_Disc_sist', index=True)
    export_agr.to_excel(writer, sheet_name='resumen', index=True)
    

    writer.save()

    Widget(my_frame3,"gray77", 15, 1, 150, 165).letra('Exportado')



#Funcion asistencia
def asistencia():
    global asist_direct_2, asist_direct_3
    import_file_path = filedialog.askopenfilename()
    asist = pd.read_excel(import_file_path,sheet_name='Base')

    #Borramos todo despues de la primear coluna vacia
    asist = asist[asist['nombres'].isna().cumsum() == 0]

    asist = asist[asist["Observacion"] != "Finiquitado"]

    #Retiramos los espacios
    asist['Observacion'] = asist['Observacion'].str.strip()

    asist["A"] = np.where(asist.d_turno == "Turno 14x14 A Dia Contingencia",1,0)
    asist["B"] = np.where(asist.d_turno == "Turno 14x14 B Dia Contingencia",1,0)
    asist["C"] = np.where(asist.d_turno == "Turno 14x14 C Dia Contingencia",1,0)
    asist["D"] = np.where(asist.d_turno == "Turno 14x14 D Dia Contingencia",1,0)
    asist["C_Noche"] = np.where(asist.d_turno == "Turno 15x13 C Noche",1,0)
    asist["D_Noche"] = np.where(asist.d_turno == "Turno 15x13 D Noche",1,0)

    asist_direct = asist[asist["Clase"] == "Directo"]
    asist_indirect = asist[asist["Clase"] == "Indirecto"]
    asist_operador = asist[asist["Clase"] == "Operador-Rigger"]

    #Reemplazamos los iteems indicados según se indica
    asist_direct["Especialidad"] = asist_direct["Especialidad"].replace(["Enfierradura","Albañilería","Carpintería", "Hormigón"], "Civil")
    asist_direct = asist_direct[["Especialidad","d_cargo","A","B","C","D","C_Noche","D_Noche","Observacion"]]
    asist_direct_1 = asist_direct[["Especialidad","d_cargo","A","B","C","D","C_Noche","D_Noche"]]
    asist_direct_1 = asist_direct_1.groupby(["Especialidad","d_cargo"]).sum()

    print(asist_direct_1)

    temp = asist_direct[asist_direct["Observacion"] == "En Obra"]

    temp_noche = temp[(temp["C_Noche"] == 1) | (temp["D_Noche"] == 1 )]
    temp_dia = temp[(temp["A"] == 1) | (temp["B"] == 1 ) | (temp["C"] == 1 ) | (temp["D"] == 1 )]


    temp_dia["Obra_día"] = 1
    temp_noche["Obra_Noche"] = 1

    temp_dia = temp_dia[["Especialidad","d_cargo","Obra_día"]].groupby(["Especialidad","d_cargo"]).sum()
    temp_noche = temp_noche[["Especialidad","d_cargo","Obra_Noche"]].groupby(["Especialidad","d_cargo"]).sum()

    #asist_direct = asist_direct.groupby(["Especialidad","d_cargo"]).size().unstack(1, fill_value=0)

    #La division de una columna en multiples.
    asist_direct = asist_direct.set_index(["Especialidad","d_cargo"])['Observacion'].str.get_dummies().groupby(["Especialidad","d_cargo"]).sum()

    asist_direct_2 = pd.concat([asist_direct,temp_dia,temp_noche,asist_direct_1],axis=1)

    print(asist_direct_2)

    asist_direct_3 = asist_direct_2.copy()


    asist_direct_3 = asist_direct_3.reset_index(level=1, drop=True)
    asist_direct_3 = asist_direct_3.groupby(["Especialidad"]).sum()


    Widget(my_frame4,"gray77", 1, 1, 150, 5).letra('OK')
def export_asis():
    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'Asistencia.xlsx')  # Creamos una excel y le indicamos la ruta

    # Exportar Steel
    asist_direct_2.to_excel(writer, sheet_name='Det_Disc_sist', index=True)
    asist_direct_3.to_excel(writer, sheet_name='Agru_Disc_Sist', index=True)



    Widget(my_frame4,"gray77", 1, 1, 150, 40).letra('OK')

    writer.save()
root = Tk()
root.title('Control Panel')
root.geometry("380x270")

my_notebook = ttk.Notebook(root)
my_notebook.pack()

my_frame1 = Frame(my_notebook,width=500,height=500,bg="gray77")
my_frame2 = Frame(my_notebook,width=500,height=500,bg="gray77")
my_frame3 = Frame(my_notebook,width=500,height=500,bg="gray")
my_frame4 = Frame(my_notebook,width=500,height=500,bg="gray")

my_frame1.pack(fill = "both", expand=1)
my_frame2.pack(fill="both", expand=1)
my_frame3.pack(fill="both", expand=1)
my_frame4.pack(fill="both", expand=1)

my_notebook.add(my_frame1,text="PB Steel")
my_notebook.add(my_frame2,text="Avan Pte")
my_notebook.add(my_frame3,text="System")
my_notebook.add(my_frame4,text="Asitencia")

#Frame 1
Widget(my_frame1,"gray56", 15, 2, 250, 5).boton('Importar Master',PB_import_master)
Widget(my_frame1,"gray56", 15, 2, 250, 45).boton('Importar OTEC',PB_import_otec)
Widget(my_frame1,"gray56", 15, 2, 250, 85).boton('Exportar PBI',PB_export)


#Frame 2
Widget(my_frame2,"gray56", 15, 1, 250, 5).boton('Importar Master',pte_import_master)
Widget(my_frame2,"gray56", 15, 1, 250, 45).boton('Importar Report_pte',pte_import_pte)
Widget(my_frame2,"gray56", 15, 1, 250, 85).boton('Procesar',pte_ejec)
Widget(my_frame2,"gray56", 15, 1, 250, 125).boton('Exportar Report pte',pte_export)
Widget(my_frame2,"gray56", 15, 1, 250, 165).boton('Exportar avance',pte_export_tekla)


#Frame3
Widget(my_frame3,"gray56", 15, 1, 250, 5).boton('Importar Piping',sist_import_piping)
Widget(my_frame3,"gray56", 15, 1, 250, 40).boton('Importar Mec',sist_import_mec)
Widget(my_frame3,"gray56", 15, 1, 250, 75).boton('Importar Elect',sist_import_elect)
Widget(my_frame3,"gray56", 15, 1, 250, 110).boton('Import Steel&OOCC',sist_import_steel)

Widget(my_frame3,"gray56", 15, 1, 250, 145).boton('Import Listado Sist',sist_import_list)
Widget(my_frame3,"gray56", 15, 1, 250, 180).boton('Export Report bi',sist_export_report_bi)
Widget(my_frame3,"gray56", 15, 1, 250, 215).boton('Export Report gral',sist_export_report)


#Frame4
Widget(my_frame4,"gray56",15,1,250,5).boton("Importar asistencia",asistencia)
Widget(my_frame4,"gray56",15,1,250,40).boton("Exportar asistencia",export_asis)

# Widget(my_frame3,"gray56", 15, 1, 250, 45).boton('Export PBI',sist_export_pbi)

root.mainloop()