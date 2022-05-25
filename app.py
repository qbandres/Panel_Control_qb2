import pandas as pd

df =  pd.read_excel('saldo_steel.xlsx')

df =df[['ESP','Peso_Total','Mount_Peso','Mount_Peso_saldo','Torq_Peso','Torq_Peso_saldo','total_hh','Mount_HH','Mount_HH_saldo','Torq_HH','Torq_HH_saldo']]
df['Peso_Total'] = df.Peso_Total/1000
df['Mount_Peso'] = df.Mount_Peso/1000
df['Mount_Peso_saldo'] = df.Mount_Peso_saldo/1000
df['Torq_Peso'] = df.Torq_Peso/1000
df['Torq_Peso_saldo'] = df.Torq_Peso_saldo/1000
df['Torq_Peso_saldo'] = df.Torq_Peso_saldo/1000

dg = df.groupby(['ESP']).sum()

dg.to_excel('Distr.xlsx')
print(df.columns)
print(dg)

