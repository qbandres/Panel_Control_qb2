import pandas as pd


df = pd.read_excel('otec_steel.xlsx',sheet_name='Estructuras',skiprows=6)

df = df.iloc[:38505]    #Filtramos solo las columnas necesarias

print(df)
print(df.columns)

df=df[['ESP','Peso Total (Kg)','Traslado.1','Pre-Armado.1','Montaje.1', 'Nivelación, Soldadura & Torque.1',
        'Punch List','total hh','traslado hh','prearm hh','montaje hh','niv, sold torq hh','punch list hh']]
dg = df.groupby('ESP').sum()

print(df)
print(df.columns)
print(dg)
print(dg.columns)