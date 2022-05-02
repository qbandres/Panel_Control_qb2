import pandas as pd


df = pd.read_excel('temp.xlsx')
def f(x):
    didx = pd.date_range(x['Actual'], x['FIN'])
    return pd.Series([didx.isocalendar().week.values, 
                      didx.strftime('%a').values], 
                      index=['Weeks', 'Days'])

df[['Weeks', 'Days']] = df.apply(f, axis=1)

print(df)