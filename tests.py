import pandas as pd
from datetime import date, timedelta
import datetime as dt 


df = pd.read_excel('temp.xlsx')

print(df.iloc[:,0])

def distri(df):
    n=0

    for i in df.iloc[:,0]:
        
        ft = [df.iloc[n,16] + timedelta(days=d) for d in range((df.iloc[n,11] - df.iloc[n,16]).days + 1)]  # CREAMOS LA LISTA DE FECHAS
        dft = pd.DataFrame({'Fecha': ft})
        n=n+1
        dft['HH_dia'] = df.HH_SALDO/len(dft)
        
        print(dft)




