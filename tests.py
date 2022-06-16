import pandas as pd
import numpy as np

df_1 = pd.DataFrame({"id": [1,2,3,4,5],
                    "fruta": ["manzana", "pera", "platano", "naranja", "aguacate"],
                    "precio": [0.20, 0.45, 0.15, 0.12, 0.62]})
print(df_1)

df_2 = pd.DataFrame({"id":[5,4,3,2,1],
                     "stock": [10, 20, 25, 12, 40]})

print(df_2)

df_3 = pd.DataFrame({"id":[4,2,5,1,3],
                     "ventas_totales":[3, 5, 2, 3, 6],
                     "ingresos_ventas": [120, 110, 64,44, 147]})

print(df_3)

df_1.merge(df_2, on="id", how="left")

print(df_1)


def conditions(x):
        if x == 'manzana'| x =='pera'| x=='platano':
            return '0310-NZC-001_1'
        elif x == 'naranja':
            return 0.9
        else:
            return 10

func = np.vectorize(conditions)
energy_class = func(df_1["fruta"])
    
df_1["aprox"] = energy_class

print(df_1)


