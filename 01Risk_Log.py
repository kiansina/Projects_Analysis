import pandas as pd
import numpy as np
df = pd.read_excel(r"C:\Users\sina.kian\Desktop\Quick report format V2\update\SRC_Questionario Base_Integrale2020_ABACO.xlsx",sheet_name=None)
df10=df['Registro_Rischi']
df10.columns=['nan','Tipologia di rischio (Strategica)','Tipologia di rischio (Ferma)','Impatto1','Causa','Eventi specifici rilenvati per l’azienda','Per trattamento assicurativo','Puro/Speculativo','Top/ Non Top','Esterno/Interno','Presente','Evento','Note','Probabilità ','Impatto2', 'nan', 'nan', 'nan', 'nan', 'nan']
df10=df10[1:]
df10.index=range(0,len(df10))
df10=df10[['Tipologia di rischio (Strategica)','Tipologia di rischio (Ferma)','Causa','Evento','Probabilità ','Impatto2']]
df10['Tipologia di rischio (Ferma)'].fillna(method="ffill", inplace=True)
df10.dropna(subset=['Impatto2'],inplace=True)
df10.sort_values(by=['Impatto2','Probabilità '],inplace=True,ascending=False)
df10['Causa_ID']=range(1,len(df10)+1)
dR=pd.DataFrame()
dR['Impatto']=[1,2,3,4,5]*5
dR['Probabilità']=[0]*len(dR)
for j in dR.index:
    dR['Probabilità'][j]=(j//5)+1

dR['Causa']=[None]*len(dR)
dR['Causa_ID']=[None]*len(dR)
dR['Tipologia di rischio (Strategica)']=[None]*len(dR)
dR['Tipologia di rischio (Ferma)']=[None]*len(dR)
dR['Evento']=[None]*len(dR)
for R in dR.index:
    dR['Causa'][R]=[df10['Causa'][x] for x in df10.index if (df10['Probabilità '][x]==dR['Probabilità'][R]) and (df10['Impatto2'][x]==dR['Impatto'][R])]
    dR['Causa_ID'][R]=[df10['Causa_ID'][x] for x in df10.index if (df10['Probabilità '][x]==dR['Probabilità'][R]) and (df10['Impatto2'][x]==dR['Impatto'][R])]
    dR['Tipologia di rischio (Strategica)'][R]=[df10['Tipologia di rischio (Strategica)'][x] for x in df10.index if (df10['Probabilità '][x]==dR['Probabilità'][R]) and (df10['Impatto2'][x]==dR['Impatto'][R])]
    dR['Tipologia di rischio (Ferma)'][R]=[df10['Tipologia di rischio (Ferma)'][x] for x in df10.index if (df10['Probabilità '][x]==dR['Probabilità'][R]) and (df10['Impatto2'][x]==dR['Impatto'][R])]
    dR['Evento'][R]=[df10['Evento'][x] for x in df10.index if (df10['Probabilità '][x]==dR['Probabilità'][R]) and (df10['Impatto2'][x]==dR['Impatto'][R])]


#dRR=pd.DataFrame()
dR=dR[['Impatto','Probabilità','Causa_ID','Causa','Tipologia di rischio (Strategica)','Tipologia di rischio (Ferma)','Evento']]


def isnan(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return False

for i in dR.index:
    ii=[]
    for j in dR['Tipologia di rischio (Strategica)'][i]:
        if j not in ii:
            if isnan(j)==False:
                ii.append(j)
                dR['Tipologia di rischio (Strategica)'][i]=ii
    ii=[]
    for j in dR['Tipologia di rischio (Ferma)'][i]:
        if j not in ii:
            if isnan(j)==False:
                ii.append(j)
                dR['Tipologia di rischio (Ferma)'][i]=ii
    ii=[]
    for j in dR['Evento'][i]:
        if j not in ii:
            if isnan(j)==False:
                ii.append(j)
                dR['Evento'][i]=ii

dR=dR.sort_values(by=['Impatto','Probabilità'])
file_name='An_Registro_Rischio.xlsx'
dR.to_excel(file_name)

df10=df10[['Tipologia di rischio (Strategica)','Tipologia di rischio (Ferma)','Causa','Causa_ID','Probabilità ', 'Impatto2','Evento']]
file_name='Elenco_di_Causa.xlsx'
df10.to_excel(file_name)
