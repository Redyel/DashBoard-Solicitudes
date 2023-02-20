import pandas as pd
import openpyxl
import numpy as np
from datetime import datetime

#Cargar el informe de Q10 y asignarlo a un dataframe
aplazamientos = pd.read_excel("./Verficacion_Academica.xlsx",engine='openpyxl')
df = pd.DataFrame(aplazamientos)

#Cambiar la fecha de interés en la extracción de datos
df['Hora'] = df['Fecha Seguimiento']
df['Fecha de Radicado'] = pd.to_datetime(df['Fecha de Radicado'], dayfirst= True).dt.date
df['Fecha Seguimiento'] = pd.to_datetime(df['Fecha Seguimiento'], dayfirst= True).dt.date

#Validar el dataframe que no contenga espacios vacíos
cola = df['Radicado'].count()
cola = cola-1
df['Radicado'] = df['Radicado'].fillna(0)
if df.loc[cola,'Radicado'] == 0:
    df.drop([cola], axis=0, inplace=True)
    df["Radicado"]= df["Radicado"].astype(float).astype(int).astype(str)
else:
    df["Radicado"]= df["Radicado"].astype(float).astype(int).astype(str)

#Concatenar el radicado y la fecha para generar un ordenamiento
df['FechaOrdenar'] = df['Fecha Seguimiento'].astype(str)
df['Ordenar'] = (df['Radicado'] +"-"+ df['FechaOrdenar'])
df.sort_values('Ordenar', inplace=True)

#Crear un nuevo DF con los datos necesarios
ndf = df.assign(demora=0)[['Radicado','Fecha de Radicado','Fecha Seguimiento','Dependencia Seguimiento','demora','Estado Solicitud','Usuario Seguimiento','Estado Seguimiento','Hora']]

#Crear el nuevo index
ndf['Index'] = np.arange(len(ndf))
ndf = ndf.set_index('Index')

#Asignar la fecha del día trabajado para los casos no gestionados
variable = (range(ndf['Radicado'].count()-1))
x = 0
y = 0
hoy = datetime.now()
format= hoy.strftime('%Y-%m-%d')

#Calcular los días de gestión de una solicitud
if(ndf.loc[0,'Estado Solicitud']) == "Pendiente":
    ndf.loc[0,'Fecha Seguimiento'] = format
    ndf['Fecha Seguimiento'] = pd.to_datetime(ndf['Fecha Seguimiento'], dayfirst=True).dt.date
    ndf.loc[0,'demora'] = (ndf.loc[0,'Fecha Seguimiento'])-(ndf.loc[0,'Fecha de Radicado'])
else:
    ndf.loc[:,'demora']=(((ndf.loc[0,'Fecha Seguimiento']))-((ndf.loc[0,'Fecha de Radicado'])))

for x in variable:
    if(ndf.loc[x+1,'Estado Solicitud']) == "Pendiente":
        ndf.loc[x+1,'Fecha Seguimiento'] = format
        ndf['Fecha Seguimiento'] = pd.to_datetime(ndf['Fecha Seguimiento'], dayfirst=True).dt.date
        ndf.loc[x+1,'demora'] = (ndf.loc[x+1,'Fecha Seguimiento'])-(ndf.loc[x+1,'Fecha de Radicado'])
    else:
        if(ndf.loc[x+1,'Radicado']) != (ndf.loc[x,'Radicado']):
            ndf.loc[x+1,'demora'] = (ndf.loc[x+1,'Fecha Seguimiento'])-(ndf.loc[x+1,'Fecha de Radicado'])
        else:
            if (ndf.loc[x+1,'Radicado']) == (ndf.loc[x,'Radicado']):
                ndf.loc[x+1,'demora'] = ((ndf.loc[x+1,'Fecha Seguimiento'])-(ndf.loc[x,'Fecha Seguimiento']))
                x+= 1

#Crear un nuevo DF con los días de gestión calculados e imprimir en CSV
xndf = ndf[['Radicado','Fecha de Radicado','Fecha Seguimiento','Dependencia Seguimiento','demora','Estado Solicitud','Usuario Seguimiento','Estado Seguimiento','Hora']]
xndf['demora'] = pd.to_timedelta(xndf['demora']).dt.days
print(xndf)
xndf.to_csv('Ok_Verficacion_Academica.csv',encoding='latin1')