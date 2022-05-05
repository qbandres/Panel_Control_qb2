import pandas as pd
from datetime import date, timedelta
import datetime as dt 


df = pd.read_excel('temp.xlsx')

print(df.iloc[:,0])

class distribucion:
    def __init__(self,df):
        self.df=df
    def distri(self):
        n=0

        for i in self.df.iloc[:,0]:
            
            ft = [self.df.iloc[n,16] + timedelta(days=d) for d in range((self.df.iloc[n,11] - self.df.iloc[n,16]).days + 1)]  # CREAMOS LA LISTA DE FECHAS
            dft = pd.DataFrame({'Fecha': ft})
            n=n+1

 
            dft['HH_dia'] = self.df.iloc[].HH_SALDO/len(dft)

            print(dft)




