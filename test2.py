import pandas as pd
import numpy as np
 

data = {'Synchronization ID': ['A', 'A.1', 'B', 'B.1', 'B.2'],'tarea': ['Primera','Segunda', 'Tercera','Cuarat' , 'Quinta'], 
'display': ['A', 'A.1', 'B', 'B.1', 'B.2']}
df = pd.DataFrame(data)


df.to_excel('naviswork.xlsx')

print(df)
