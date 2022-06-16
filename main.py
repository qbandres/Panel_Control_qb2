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

    # esp_comp = pd.concat([dg_otec,dfv_g],axis=1)



    # Exportar Steel
    dfv.to_excel(writer, sheet_name='ST_Gan', index=True)
    df_base.to_excel(writer, sheet_name='ST_Base', index=True)
    # df_otec.to_excel(writer, sheet_name='ST_Otec', index=True)
    # esp_comp.to_excel(writer, sheet_name='Comparacion_ESP', index=True)

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
                        'AVANCE FINAL','HH TOTAL\nML','HH\nAVANCE']]

    pip_sist_men = pip_sist_2[['SUB \nSISTEMA','Codigo Fluido','ISOMETRICOS','Metros Aislados','Diametro','Metros cañerias',
                        'AVANCE TOTAL\nAG','HH TOTAL','HH\nAVANCE']]
    
    pip_sist_valv = pip_sist_3[['SUB-SISTEMA','Codigo Fluido','Isometrico','Diametro','Cantidad',
                         'AVANCE TOTAL\nAG','HH TOTAL','HH\nAVANCE']]
    pip_sist_soport = pip_sist_4[['Isometrico','Diametro','Peso soporte total KG','HH\nAVANCE','HH TOTAL\nML','AVANCE KG\nAG']]

    print(pip_sist_gen.columns)
    print(pip_sist_gen.columns)
    print(pip_sist_gen.columns)
    print(pip_sist_gen.columns)


    #CUBICACION GENERAL
    pip_sist_gen.rename(columns={'COD': 'QUIEBRE_OT','SUB-SISTEMA': 'SUBSISTEMA','HH TOTAL\nML': 'HH_Tot','Metros cañerias fore cast 11': 'Cant_Tot','Diametro':'DIAMETRO','AVANCE FINAL':'Cant_Avan',
                                'HH\nAVANCE':'HH_Avan','ISOMETRICO':'TAG','UG/AG/OR':'Niv','Codigo Fluido':'Fluid'},
                       inplace=True)
    
    print(pip_sist_gen.columns)

    #INFO TEMPORAL PARA EXTRAR EL SUBSISTEMA
    iso_sub = pip_sist_gen[['TAG','SUBSISTEMA','Fluid']]
    iso_sub['ISO'] = iso_sub['TAG'].str[:13]
    del iso_sub['TAG']

    iso_sub = iso_sub.groupby(['ISO']).first()
    iso_sub = iso_sub.reset_index()


    pip_sist_gen = pip_sist_gen[pip_sist_gen['QUIEBRE_OT'] != 'FA']

    pip_sist_gen = pip_sist_gen.fillna(0)   

    pip_sist_gen['HH_Saldo'] = pip_sist_gen['HH_Tot'].subtract(pip_sist_gen["HH_Avan"])
    pip_sist_gen.insert(9, 'Cant_Sal', pip_sist_gen['Cant_Tot'].subtract(pip_sist_gen["Cant_Avan"]))

    pip_sist_gen['OTEC'] = 'General'
    pip_sist_gen['Und'] = 'ml'
    pip_sist_gen['ISO'] = pip_sist_gen['TAG'].str[:13]

    pip_sist_gen.groupby(['SUBSISTEMA','TAG']).sum().to_excel('pipingdesglose.xlsx')

    pip_sist_gen = pip_sist_gen[pip_sist_gen['QUIEBRE_OT'] != 0]

    #Se añade columnas adicionales HHConst & HH Punch




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
    pip_sist_soport.rename(columns={'Peso soporte total KG': 'Cant_Tot','Diametro':'DIAMETRO','AVANCE KG\nAG':'Cant_Avan','Isometrico':'TAG','HH\nAVANCE':'HH_Avan','HH TOTAL\nML': 'HH_Tot'},inplace=True)

    pip_sist_soport['ISO'] = pip_sist_soport['TAG'].str[:13]

    pip_sist_soport = pip_sist_soport.groupby(['ISO']).sum()

    pip_sist_soport = pip_sist_soport.reset_index()
    pip_sist_soport['HH_Saldo'] = pip_sist_soport['HH_Tot'].subtract(pip_sist_soport["HH_Avan"])



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

    print(pip_tot)

    pip_tot['disc'] = 'PIP'
    pip_tot_res = pip_tot.groupby(['QUIEBRE_OT']).sum()

    pip_conc = pip_tot[['QUIEBRE_OT', 'SUBSISTEMA','TAG','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]

    pip_conc = pip_conc[pip_conc['SUBSISTEMA'] != 0 ]

    pip_conc['HH_Tot_C'] = pip_conc.HH_Tot*0.95
    pip_conc['HH_Sal_C'] = np.where(pip_conc.HH_Avan < pip_conc.HH_Tot_C,pip_conc.HH_Tot_C - pip_conc.HH_Avan,0)
    pip_conc['HH_Tot_P'] = pip_conc.HH_Tot*0.05
    pip_conc['HH_Sal_P'] = np.where(pip_conc.HH_Avan <= pip_conc.HH_Tot_C,pip_conc.HH_Tot_P,pip_conc.HH_Tot - pip_conc.HH_Avan)

    pip_conc = pip_conc.fillna(0)
    pip_conc = pip_conc[pip_conc['HH_Tot'] != 0]


    Widget(my_frame3,"gray77", 15, 1, 150, 5).letra('Importado')
def sist_import_mec():
    global mec_sist, mec_conc
    import_file_path = filedialog.askopenfilename()
    mec_sist = pd.read_excel(import_file_path,sheet_name='Base Datos',skiprows=6)
    print(mec_sist.columns)


    mec_sist = mec_sist[['TAG','SUBSIST','UND','HyH UND Rev. 12','HH SALDO Rev. 12','Disciplina','Area']]
    mec_sist.rename(columns={'SUBSIST': 'SUBSISTEMA','HyH UND Rev. 12': 'HH_Tot','HH SALDO Rev. 12': 'HH_Saldo','UND':'Und'},
                       inplace=True)

    mec_sist = mec_sist[mec_sist['Area'] == 310]
    del mec_sist['Area']

    mec_sist = mec_sist.dropna(subset=['Und'])


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
    energy_class = func(mec_sist["Und"])

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

    print(mec_sist)
    print(mec_sist.columns)


    mec_conc = mec_sist[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]

    mec_conc['SUBSISTEMA'] = np.where(mec_conc['SUBSISTEMA'] == 0, 'POR DEFINIR OTEC',mec_conc['SUBSISTEMA'])
    

    print(mec_conc)


    mec_conc['HH_Tot_C'] = mec_conc.HH_Tot*0.95
    mec_conc['HH_Sal_C'] = np.where(mec_conc.HH_Avan < mec_conc.HH_Tot_C,mec_conc.HH_Tot_C - mec_conc.HH_Avan,0)
    mec_conc['HH_Tot_P'] = mec_conc.HH_Tot*0.05
    mec_conc['HH_Sal_P'] = np.where(mec_conc.HH_Avan <= mec_conc.HH_Tot_C,mec_conc.HH_Tot_P,mec_conc.HH_Tot - mec_conc.HH_Avan)

    mec_conc = mec_conc.fillna(0)
    mec_conc = mec_conc[mec_conc['HH_Tot'] != 0]

    print(mec_conc[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','disc']].groupby(['disc']).sum())


    Widget(my_frame3,"gray77", 15, 1, 150, 40).letra('Importado')
def sist_import_elect():
    global elec_conc_total
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
    elect_sist_cab['disc'] = 'ELEC'
   

    elec_con_cab = elect_sist_cab[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    elec_con_cab = elec_con_cab[elec_con_cab['SUBSISTEMA'] != 'Eliminado' ]

    elec_con_cab['HH_Tot_C'] = elec_con_cab.HH_Tot*0.7
    elec_con_cab['HH_Sal_C'] = np.where(elec_con_cab.HH_Avan < elec_con_cab.HH_Tot_C,elec_con_cab.HH_Tot_C - elec_con_cab.HH_Avan,0)
    elec_con_cab['HH_Tot_P'] = elec_con_cab.HH_Tot*0.3
    elec_con_cab['HH_Sal_P'] = np.where(elec_con_cab.HH_Avan <= elec_con_cab.HH_Tot_C,elec_con_cab.HH_Tot_P,elec_con_cab.HH_Tot - elec_con_cab.HH_Avan)


    #Alumbrado

    elect_sist_alu=elect_sist_alu.iloc[:, lambda elect_sist_alu: [8,21,23,28,32,46,48,49]]
    elect_sist_alu.columns = ['TAG','Und','Cant_Tot','Tendido','Cant_Avan','QUIEBRE_OT','HH_Tot','HH_Avan']

    # elect_sist_alu['QUIEBRE_OT'] = elect_sist_alu['QUIEBRE_OT'].fillna(value='Por definir')
    elect_sist_alu = elect_sist_alu.dropna(subset=['QUIEBRE_OT'])
    elect_sist_alu = elect_sist_alu.dropna(how='all')

    elect_sist_alu['HH_Saldo'] = elect_sist_alu['HH_Tot'].subtract(elect_sist_alu["HH_Avan"])
    elect_sist_alu['Cant_Sal'] = elect_sist_alu['Cant_Tot'].subtract(elect_sist_alu["Cant_Avan"])
    elect_sist_alu['OTEC'] = 'ALUMBRADO'
    elect_sist_alu['disc'] = 'ELEC'
    elect_sist_alu['SUBSISTEMA'] = '0310-NZC-001'

    elec_con_al = elect_sist_alu[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]

    def conditions(x):
        if x == 'Cable Alumbrado':
            return 0.7
        elif x == 'Conduit Alumbrado':
            return 0.9
        elif x == 'Luminarias':
            return 0.7
        else:
            return 0.9

    func = np.vectorize(conditions)
    energy_class = func(elec_con_al["QUIEBRE_OT"])
    
    elec_con_al["factor"] = energy_class

    elec_con_al['HH_Tot_C'] = elec_con_al.HH_Tot*elec_con_al.factor
    elec_con_al['HH_Sal_C'] = np.where(elec_con_al.HH_Avan < elec_con_al.HH_Tot_C,elec_con_al.HH_Tot_C - elec_con_al.HH_Avan,0)
    elec_con_al['HH_Tot_P'] = elec_con_al.HH_Tot*(1-elec_con_al.factor)
    elec_con_al['HH_Sal_P'] = np.where(elec_con_al.HH_Avan <= elec_con_al.HH_Tot_C,elec_con_al.HH_Tot_P,elec_con_al.HH_Tot - elec_con_al.HH_Avan)
    del elec_con_al['factor']

    

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
    elect_sist_malla['disc'] = 'ELEC'
    elect_sist_malla['SUBSISTEMA'] = '0310-NZC-001'

    elec_con_mall = elect_sist_malla[['QUIEBRE_OT','TAG','SUBSISTEMA','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]

    elec_con_mall['HH_Tot_C'] = elec_con_mall.HH_Tot*0.9
    elec_con_mall['HH_Sal_C'] = np.where(elec_con_mall.HH_Avan < elec_con_mall.HH_Tot_C,elec_con_mall.HH_Tot_C - elec_con_mall.HH_Avan,0)
    elec_con_mall['HH_Tot_P'] = elec_con_mall.HH_Tot*0.1
    elec_con_mall['HH_Sal_P'] = np.where(elec_con_mall.HH_Avan <= elec_con_mall.HH_Tot_C,elec_con_mall.HH_Tot_P,elec_con_mall.HH_Tot - elec_con_mall.HH_Avan)

    #EPC

    elect_sist_epc=elect_sist_epc.iloc[:, lambda elect_sist_epc: [9,16,17,24,26,42,44,45]]
    elect_sist_epc.columns = ['TAG','Cant_Tot','Und','Tendido','Cant_Avan','QUIEBRE_OT','HH_Tot','HH_Avan']

    elect_sist_epc = elect_sist_epc[elect_sist_epc['QUIEBRE_OT'] != 'Descope']
    elect_sist_epc = elect_sist_epc[elect_sist_epc['QUIEBRE_OT'] != 'Eliminado']


    elect_sist_epc['HH_Saldo'] = elect_sist_epc['HH_Tot'].subtract(elect_sist_epc["HH_Avan"])
    elect_sist_epc['Cant_Sal'] = elect_sist_epc['Cant_Tot'].subtract(elect_sist_epc["Cant_Avan"])
    elect_sist_epc['OTEC'] = 'EPC'
    elect_sist_epc['disc'] = 'ELEC'
    elect_sist_epc['SUBSISTEMA'] = '0310-NZC-001'

    elec_con_epc = elect_sist_epc[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    print(elec_con_epc)

    elec_con_epc['HH_Tot_C'] = elec_con_epc.HH_Tot*0.95
    elec_con_epc['HH_Sal_C'] = np.where(elec_con_epc.HH_Avan < elec_con_epc.HH_Tot_C,elec_con_epc.HH_Tot_C - elec_con_epc.HH_Avan,0)
    elec_con_epc['HH_Tot_P'] = elec_con_epc.HH_Tot*0.05
    elec_con_epc['HH_Sal_P'] = np.where(elec_con_epc.HH_Avan <= elec_con_epc.HH_Tot_C,elec_con_epc.HH_Tot_P,elec_con_epc.HH_Tot - elec_con_epc.HH_Avan)


    #EQUIPOS
    elect_sist_equ=elect_sist_equ.iloc[:, lambda elect_sist_equ: [4,8,15,16,23,27,42,44,46,48]]
    elect_sist_equ.columns = ['area','TAG','Und','Cant_Tot','Montaje','Cant_Avan','HH_Tot','HH_Avan','SUBSISTEMA','QUIEBRE_OT']

    elect_sist_equ = elect_sist_equ[elect_sist_equ['area'] != 320]
    elect_sist_equ['QUIEBRE_OT'] = elect_sist_equ['QUIEBRE_OT'].fillna('INDEFINIDO')

    elect_sist_equ['HH_Saldo'] = elect_sist_equ['HH_Tot'].subtract(elect_sist_equ["HH_Avan"])
    elect_sist_equ['Cant_Sal'] = elect_sist_equ['Cant_Tot'].subtract(elect_sist_equ["Cant_Avan"])
    elect_sist_equ['OTEC'] = 'Equipos electricos'
    elect_sist_equ['disc'] = 'ELEC'


    elec_con_equ = elect_sist_equ[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    elec_con_equ = elec_con_equ[elec_con_equ['SUBSISTEMA'] != 'Eliminado' ]
    elec_con_equ['SUBSISTEMA'] = elec_con_equ['SUBSISTEMA'].fillna('POR DEFINIR OTEC')
    print(elec_con_equ)

    elec_con_equ['HH_Tot_C'] = elec_con_equ.HH_Tot*0.9
    elec_con_equ['HH_Sal_C'] = np.where(elec_con_equ.HH_Avan < elec_con_equ.HH_Tot_C,elec_con_equ.HH_Tot_C - elec_con_equ.HH_Avan,0)
    elec_con_equ['HH_Tot_P'] = elec_con_equ.HH_Tot*0.1
    elec_con_equ['HH_Sal_P'] = np.where(elec_con_equ.HH_Avan <= elec_con_equ.HH_Tot_C,elec_con_equ.HH_Tot_P,elec_con_equ.HH_Tot - elec_con_equ.HH_Avan)

    Widget(my_frame3,"gray77", 15, 1, 150, 75).letra('Importado')

    #inSTRUMENTOS

    elect_sist_inst=elect_sist_inst.iloc[:, lambda elect_sist_inst: [11,3,7,30,37,39,48,49,55]] 
    elect_sist_inst.columns = ['area','TAG','SUBSISTEMA','Cant_Tot','Montaje','Cant_Avan','HH_Tot','HH_Avan','QUIEBRE_OT']
    elect_sist_inst = elect_sist_inst[elect_sist_inst['area'] != 320]
    elect_sist_inst['QUIEBRE_OT'] = elect_sist_inst['QUIEBRE_OT'].fillna('INDEFINIDO')
    elect_sist_inst['HH_Saldo'] = elect_sist_inst['HH_Tot'].subtract(elect_sist_inst["HH_Avan"])
    elect_sist_inst['Cant_Sal'] = elect_sist_inst['Cant_Tot'].subtract(elect_sist_inst["Cant_Avan"])
    elect_sist_inst['OTEC'] = 'Instrumentos'
    elect_sist_inst['disc'] = 'ELEC'
    elect_sist_inst['Und'] = 'UND'

    elec_con_inst = elect_sist_inst[['QUIEBRE_OT','SUBSISTEMA','TAG','Cant_Tot','Cant_Avan','Cant_Sal','Und','HH_Tot','HH_Avan','HH_Saldo','OTEC','disc']]
    elec_con_inst = elec_con_inst[elec_con_inst['SUBSISTEMA'] != 'Falta PID con sistema']
    elec_con_inst = elec_con_inst[elec_con_inst['SUBSISTEMA'] != 'HOLD']
    elec_con_inst['SUBSISTEMA'] = elec_con_inst['SUBSISTEMA'].fillna('POR DEFINIR OTEC')

    elec_con_inst['HH_Tot_C'] = elec_con_inst.HH_Tot*0.9
    elec_con_inst['HH_Sal_C'] = np.where(elec_con_inst.HH_Avan < elec_con_inst.HH_Tot_C,elec_con_inst.HH_Tot_C - elec_con_inst.HH_Avan,0)
    elec_con_inst['HH_Tot_P'] = elec_con_inst.HH_Tot*0.1
    elec_con_inst['HH_Sal_P'] = np.where(elec_con_inst.HH_Avan <= elec_con_inst.HH_Tot_C,elec_con_inst.HH_Tot_P,elec_con_inst.HH_Tot - elec_con_inst.HH_Avan)

    elec_conc_total = pd.concat([elec_con_cab,elec_con_al,elec_con_mall,elec_con_epc,elec_con_equ,elec_con_inst],axis=0)
    elec_conc_total = elec_conc_total.fillna(0)
    elec_conc_total = elec_conc_total[elec_conc_total['HH_Tot'] != 0]

       print(elec_con_inst)

def sist_import_steel():
    global steel_conc

    import_file_path = filedialog.askopenfilename()
    steel_sist= pd.read_excel(import_file_path,sheet_name='Resumen L1 y L2',skiprows=1)

    steel_sist=steel_sist.iloc[:, lambda elect_sist_cab: [4,5,6,7,11,12,13]]
    steel_sist.columns = ['QUIEBRE_OT','Cant_Tot','Cant_Avan','Cant_Sal','HH_Tot','HH_Avan','HH_Saldo']
    steel_sist = steel_sist.iloc[:28]

    steel_sist['TAG'] = 'No aplica'

    temp = pd.DataFrame({'QUIEBRE_OT':['03101.H032', '03101.H033', '03101.H030', '03101.H031', '03101.H108', '03101.H109', '03101.H110', '03101.H111', '03101.H112', '03101.H114', '03101.H113', '03101.H115', '03101.H116', '03101.H117', '03101.H118', '03101.H061', '03301.H062', '03201.H006', '03201.N008', '03301.H011', '09411.H003', '09101.H013', '03101.N159'],
            'SUBSISTEMA':['0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_2', '0310-NZC-001_2', '0310-NZC-001_1', '0310-NZC-001_1', '0310-NZC-001_1', '0310-NZC-001_1', '0310-NZC-001_1', '0310-NZC-001_1', '0310-NZC-001_1', '0310-NZC-001_1']
            })

    steel_sist = steel_sist.merge(temp[['QUIEBRE_OT','SUBSISTEMA']], on='QUIEBRE_OT',
                    how='left')  # Buscas HH de ID y lo insertas en data Horas ganadas
    steel_sist['Und'] = 'Ton'
    steel_sist['disc'] = 'EST'
    steel_sist['OTEC'] = 'BASE'
    steel_sist = steel_sist.dropna(subset=['SUBSISTEMA'])

    steel_conc = steel_sist[['QUIEBRE_OT', 'SUBSISTEMA','TAG','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]

    print(steel_conc)
    

    steel_conc['HH_Tot_C'] = steel_conc['HH_Tot']*0.9
    steel_conc['HH_Sal_C'] = np.where(steel_conc.HH_Avan < steel_conc.HH_Tot_C,steel_conc.HH_Tot_C - steel_conc.HH_Avan,0)
    steel_conc['HH_Tot_P'] = steel_conc['HH_Tot']*0.1
    steel_conc['HH_Sal_P'] = np.where(steel_conc.HH_Avan <= steel_conc.HH_Tot_C,steel_conc.HH_Tot_P,steel_conc.HH_Tot - steel_conc.HH_Avan)
    print(steel_conc)
    Widget(my_frame3,"gray77", 15, 1, 150, 110).letra('Importado')

def sist_import_arq():

    global arq_conc

    import_file_path = filedialog.askopenfilename()
    # oocc_sist= pd.read_excel(import_file_path,sheet_name='OOCC')
    # oocc_sist['TAG'] = 'No aplica' 
    arq_sist= pd.read_excel(import_file_path,sheet_name='Resumen L1 y L2',skiprows=2)
    arq_sist=arq_sist.iloc[:, lambda elect_sist_cab: [2,3,4,5,9,10,11]]
    arq_sist.columns = ['LINEA','Cant_Tot','Cant_Avan','Cant_Sal','HH_Tot','HH_Avan','HH_Saldo']
    arq_sist['TAG'] = 'No aplica'
    arq_sist['OTEC'] = 'base_arq'
    arq_sist['QUIEBRE_OT'] =  'Instalación'
    arq_sist['disc'] =  'ARQ'
    arq_sist['Und'] =  'm2'

    def conditions(x):
        if x == 'L1':
            return '0310-NZC-001_1'
        elif x == 'L2':
            return '0310-NZC-001_2'
        else:
            return '0310-NZC-001_1'

    func = np.vectorize(conditions)
    energy_class = func(arq_sist["LINEA"])
    
    arq_sist["SUBSISTEMA"] = energy_class

    del arq_sist['LINEA']

    arq_conc = arq_sist[['QUIEBRE_OT', 'SUBSISTEMA','TAG','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]

    arq_conc['HH_Tot_C'] = arq_conc.HH_Tot*0.9
    arq_conc['HH_Sal_C'] = np.where(arq_conc.HH_Avan < arq_conc.HH_Tot_C,arq_conc.HH_Tot_C - arq_conc.HH_Avan,0)
    arq_conc['HH_Tot_P'] = arq_conc.HH_Tot*0.1
    arq_conc['HH_Sal_P'] = np.where(arq_conc.HH_Avan <= arq_conc.HH_Tot_C,arq_conc.HH_Tot_P,arq_conc.HH_Tot - arq_conc.HH_Avan)

    print(arq_conc)

    Widget(my_frame3,"gray77", 15, 1, 150, 145).letra('Importado')

def sist_import_OC():
    global oocc_conc

    import_file_path = filedialog.askopenfilename()
    oocc_sist= pd.read_excel(import_file_path,sheet_name='OOCC')
    oocc_sist['TAG'] = 'No aplica' 

        
    oocc_conc = oocc_sist[['QUIEBRE_OT', 'SUBSISTEMA','TAG','Cant_Tot', 'Cant_Avan', 'Cant_Sal','Und', 'HH_Tot', 'HH_Avan', 'HH_Saldo',
       'OTEC','disc']]

    

    oocc_conc['HH_Tot_C'] = oocc_conc.HH_Tot
    oocc_conc['HH_Sal_C'] = oocc_conc.HH_Saldo
    oocc_conc['HH_Tot_P'] = 0
    oocc_conc['HH_Sal_P'] = 0

    print(oocc_conc)


    Widget(my_frame3,"gray77", 15, 1, 150, 180).letra('Importado')

#Import Fechas
def sist_import_list():
    global df_sub,export_agr_subsist,export_agr_subsist_d,export_agr_subsist_q,export,export_agr_quiebre,export_agr_subsist_frnete

    import_file_path = filedialog.askopenfilename()
    list_sist = pd.read_excel(import_file_path,sheet_name='data')
    list_sist['SUBSISTEMA'] =list_sist['SUBSISTEMA'].str.strip()
    list_sist['FIN'] = pd.to_datetime(list_sist.FIN).dt.date
    list_sist['Actual'] = pd.to_datetime(list_sist.Actual).dt.date 
    list_sist['Ini_P'] = pd.to_datetime(list_sist.Ini_P).dt.date
    list_sist['Fin_P'] = pd.to_datetime(list_sist.Fin_P).dt.date
    

    export = pd.concat([pip_conc,mec_conc,elec_conc_total,steel_conc,oocc_conc,arq_conc],axis=0)
    

    # def conditions(x):
    #     if x == '0310-HCO-002':
    #         return '0310-HC-002'
    #     elif x > '0310-HC-015':
    #         return "0310-HA-015"
    #     elif x > '0310-HC-015':
    #         return "0310-HA-015"
    #     elif x > '0310-HC-015':
    #         return "0310-HA-015"
    #     else:
    #         return x

    # func = np.vectorize(conditions)
    # energy_class = func(export["SUBSISTEMA"])

    # export["SUBSISTEMA"] = energy_class

     
    export_agr_quiebre = export.merge(list_sist, on='SUBSISTEMA', how='outer')
    export_agr_quiebre = export_agr_quiebre.dropna(subset=['QUIEBRE_OT'])
    print(export_agr_quiebre)

    df_sub = export.merge(list_sist, on='SUBSISTEMA', how='outer')
    df_sub = df_sub[df_sub['Saldo_dias'] > 0 ]
    del df_sub['Saldo_dias']  
    del df_sub['Saldo_dias_p']
    export_agr_quiebre = export_agr_quiebre[export_agr_quiebre['ALCANCE'] != 'NO']
    df_sub = df_sub[df_sub['ALCANCE'] != 'NO']

    
    export_agr_quiebre = df_sub.groupby(['disc','OTEC','QUIEBRE_OT']).sum()
    export_agr_subsist = df_sub.groupby(['ALCANCE','LINEA','SUBSISTEMA']).sum()
    export_agr_subsist_d = df_sub.groupby(['ALCANCE','LINEA','SUBSISTEMA','disc']).sum()
    export_agr_subsist_q = df_sub.groupby(['ALCANCE','LINEA','SUBSISTEMA','disc','QUIEBRE_OT']).sum()
    export_agr_subsist_frnete = df_sub.groupby(['ALCANCE','Frente','SUBSISTEMA']).sum()
    export_agr_quiebre['Percent'] = 1-export_agr_quiebre.HH_Saldo/export_agr_quiebre.HH_Tot
    export_agr_subsist['Percent'] = 1-export_agr_subsist.HH_Saldo/export_agr_subsist.HH_Tot
    export_agr_subsist_d['Percent'] = 1-export_agr_subsist_d.HH_Saldo/export_agr_subsist_d.HH_Tot
    export_agr_subsist_q['Percent'] = 1-export_agr_subsist_q.HH_Saldo/export_agr_subsist_q.HH_Tot
    export_agr_subsist_frnete['Percent'] = 1-export_agr_subsist_frnete.HH_Saldo/export_agr_subsist_frnete.HH_Tot

    # export_agr_subsist_q=export_agr_subsist_q.reset_index()
    export_agr_subsist_d=export_agr_subsist_d.reset_index()
    export_agr_subsist_frnete=export_agr_subsist_frnete.reset_index()

    df_sub = df_sub.groupby(['SUBSISTEMA','disc','QUIEBRE_OT']).sum()

    df_sub=df_sub.reset_index()

    df_sub = df_sub.merge(list_sist, on='SUBSISTEMA', how='outer')
    df_sub = df_sub.dropna(subset=['HH_Tot'])
    # df_sub = df_sub[df_sub['ALCANCE'] != 'NO']

    print(df_sub)
    print(df_sub.columns)
    print(export_agr_quiebre)
    print(export_agr_subsist_q)


    Widget(my_frame3,"gray77", 15, 1, 150, 215).letra('Importado')

#Exportar Data
def sist_export_data():

    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'data_sistemas_qb2.xlsx')  # Creamos una excel y le indicamos la ruta

    export.to_excel(writer, sheet_name='Detalle', index=True)
    export_agr_quiebre.to_excel(writer, sheet_name='QUIEBRE', index=True)
    export_agr_subsist.to_excel(writer, sheet_name='SUBSISTEMA', index=True)
    export_agr_subsist_d.to_excel(writer, sheet_name='SUBSISTEMA_dis', index=True)
    export_agr_subsist_q.to_excel(writer, sheet_name='SUBSISTEMA_quie', index=True)
    export_agr_subsist_frnete.to_excel(writer, sheet_name='Frente', index=True)

    Widget(my_frame3,"gray77", 1, 1, 150, 250).letra('OK')

    writer.save()

#Power Bi
def sist_export_report_bi():


    # df_sub['HH_Sem'] = df_sub.HH_Saldo/(df_sub['Saldo_dias'])*7
    # df_sub['Cant_Sem'] = df_sub.Cant_Sal/(df_sub['Saldo_dias'])*7

    n=0
    m=0

    titulos_1 = ['SUBSISTEMA', 'disc', 'QUIEBRE_OT', 'Cant_Tot', 'Cant_Avan', 'Cant_Sal',
       'HH_Tot', 'HH_Avan', 'HH_Saldo', 'HH_Tot_C', 'HH_Sal_C', 'HH_Tot_P',
       'HH_Sal_P', 'CODE', 'AREA', 'DESCRIP', 'CONTRATO', 'TIPO', 'ALCANCE',
       'LINEA', 'Actual', 'FIN', 'Saldo_dias', 'Frente', 'Ini_P', 'Fin_P',
       'Saldo_dias_p']

    # titulos_2 =  ['SUBSISTEMA', 'disc', 'QUIEBRE_OT', 'Cant_Tot', 'Cant_Avan', 'Cant_Sal',
    #    'HH_Tot', 'HH_Avan', 'HH_Saldo', 'HH_Tot_C', 'HH_Sal_C', 'HH_Tot_P',
    #    'HH_Sal_P', 'CODE', 'AREA', 'DESCRIP', 'CONTRATO', 'TIPO', 'ALCANCE',
    #    'LINEA', 'Actual', 'FIN', 'Saldo_dias', 'Frente', 'Ini_P', 'Fin_P',
    #    'Saldo_dias_p']

    df_temp_c = pd.DataFrame(columns=titulos_1)
    df_temp_p = pd.DataFrame(columns=titulos_1)
 
    for i in df_sub.index:
        print(i)
        ftc = [df_sub.iloc[n,20] + timedelta(days=d) for d in range((df_sub.iloc[n,21] - df_sub.iloc[n,20]).days + 1)]  # CREAMOS LA LISTA DE FECHAS
        dftc = pd.DataFrame({'FECHA': ftc})

        dftc['HH_Dia'] = df_sub.iloc[n,10]/len(dftc)
        dftc['Cant_Dia'] = df_sub.iloc[n,5]/len(dftc)
        dftc['HH_Final'] = df_sub.iloc[n,9]/len(dftc)
        dftc['Cant_Final'] = df_sub.iloc[n,3]/len(dftc)
        dftvc = df_sub.iloc[[i]]
        dftvc['FECHA'] = ftc[0]
        dftvc.set_index('FECHA',inplace =True )
        print(dftvc)
        dftc.set_index('FECHA',inplace =True)
        resc = pd.concat([dftvc,dftc],axis=1)
        resc = resc.fillna(method='ffill')   
        print(resc)
        df_temp_c = pd.concat([df_temp_c,resc],axis=0)
        df_temp_c['etapa'] = 'Constr'
        print(df_temp_c)
        print(n)
        n=n+1

    for j in df_sub.index:
        print(j)
        ftp = [df_sub.iloc[m,24] + timedelta(days=d) for d in range((df_sub.iloc[m,25] - df_sub.iloc[m,24]).days + 1)]  # CREAMOS LA LISTA DE FECHAS
        dftp = pd.DataFrame({'FECHA': ftp})

        dftp['HH_Dia'] = df_sub.iloc[m,12]/len(dftp)
        dftp['Cant_Dia'] = 0
        dftp['HH_Final'] = df_sub.iloc[m,11]/len(dftp)
        dftp['Cant_Final'] = 0
        dftvp = df_sub.iloc[[j]]
        dftvp['FECHA'] = ftp[0]
        dftvp.set_index('FECHA',inplace =True)
        print(dftvp)
        dftp.set_index('FECHA',inplace =True)
        resp = pd.concat([dftvp,dftp],axis=1)
        resp = resp.fillna(method='ffill')   
        print(resp)
        df_temp_p = pd.concat([df_temp_p,resp],axis=0)
        df_temp_p['etapa'] = 'Punch'
        print(df_temp_p)
        print(m)
        m=m+1


    df_temp = pd.concat([df_temp_c,df_temp_p],axis=0)

    del df_temp['Cant_Avan']
    del df_temp['Cant_Tot']
    del df_temp['Cant_Sal']
    del df_temp['HH_Tot']
    del df_temp['HH_Saldo']
    del df_temp['HH_Avan']
    del df_temp['HH_Tot_C']
    del df_temp['HH_Sal_C']
    del df_temp['HH_Tot_P']
    del df_temp['HH_Sal_P']
    del df_temp['CODE']
    del df_temp['Ini_P']
    del df_temp['Fin_P']
    del df_temp['Actual']
    del df_temp['FIN']
    del df_temp['Saldo_dias_p']
    del df_temp['Saldo_dias']

    # del df_temp['HH_Sem']

    df_temp.reset_index(level=0, inplace=True)
    df_temp.rename(columns={"index": "Fecha"},inplace=True)
    df_temp['FECHA']=pd.to_datetime(df_temp.Fecha).dt.date   
    df_temp = Semana(df_temp).split()  

    df_temp['disc'] = np.where(df_temp.disc.isnull(),"NA",df_temp['disc'])
    df_temp = df_temp[df_temp['ALCANCE'] != 'NO']

    del df_temp['Fecha']
    # del df_temp[df_temp.columns[0]]
        
    export_file = filedialog.askdirectory()  # Buscamos el directorio para guardar
    writer = pd.ExcelWriter(export_file + '/' + 'PB_Sistemas_qb2.xlsx')  # Creamos una excel y le indicamos la ruta

    df_temp.to_excel(writer, sheet_name='PowerBi', index=False)
    df_sub.to_excel(writer, sheet_name='data', index=False)

    Widget(my_frame3,"gray77", 1, 1, 150, 285).letra('OK')

    writer.save()

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
root.geometry("380x420")

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
Widget(my_frame3,"gray56", 15, 1, 250, 110).boton('Import Steel',sist_import_steel)
Widget(my_frame3,"gray56", 15, 1, 250, 145).boton('Import ARQ',sist_import_arq)
Widget(my_frame3,"gray56", 15, 1, 250, 180).boton('Import OOCC',sist_import_OC)



Widget(my_frame3,"gray56", 15, 1, 250, 215).boton('Import Listado Sist',sist_import_list)
Widget(my_frame3,"gray56", 15, 1, 250, 250).boton('Exportar Data',sist_export_data)
Widget(my_frame3,"gray56", 15, 1, 250, 285).boton('Export Report bi',sist_export_report_bi)


#Frame4
Widget(my_frame4,"gray56",15,1,250,5).boton("Importar asistencia",asistencia)
Widget(my_frame4,"gray56",15,1,250,40).boton("Exportar asistencia",export_asis)

# Widget(my_frame3,"gray56", 15, 1, 250, 45).boton('Export PBI',sist_export_pbi)

root.mainloop()