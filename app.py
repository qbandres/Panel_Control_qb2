import pandas as pd

df =  pd.read_excel('temp_subsitem.xlsx')

quiebre = df['QUIEBRE_OT'].tolist()
subsistem = df['SUBSISTEMA'].tolist()

print(quiebre)
print(subsistem)
