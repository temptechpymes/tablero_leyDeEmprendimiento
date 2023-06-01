#importar librerias 
import pandas as pd
from datetime import datetime as date  
import numpy as np
import math
import os 
import win32com.client as client
from datetime import datetime
import openpyxl


#  utilizando la biblioteca pandas para leer un archivo Excel y eliminar columnas
#  específicas del DataFrame.
data = pd.read_excel('reporte.xlsx')
#Elimina columnas
data.pop('firstname') 
data.pop('lastname')
data.pop('timecreated')
data.pop('id')


#  Itera sobre los índices de la columna 'email' y realiza transformaciones en los datos
for i in range(0, len(data['email'])):
    data['email'][i] = data['email'][i].strip().lower()
    if(str(data['idnumber'][i]) == 'nan'):
        data['idnumber'][i] = ''
    else:
        data['idnumber'][i] = str(data['idnumber'][i]).strip().replace('.', '')
    if data['days'][i] != 'No ha ingresado' and date.strptime(data['days'][i], '%Y-%m-%d %H:%M:%S') > date(day=27, month=9, year=2022):
        data['days'][i] = date.strptime(data['days'][i], '%Y-%m-%d %H:%M:%S').date().strftime('%d/%m/%Y')
    else:
        if data['days'][i] != 'No ha ingresado':
            if date.strptime(data['days'][i], '%Y-%m-%d %H:%M:%S') < date(day=28, month=9, year=2022):
                data['days'][i] = 'N/A'
        else:
            data['days'][i] = 'NO'
    print(i)

with pd.ExcelWriter('reporteFormateado.xlsx') as writer:
    data.to_excel(writer, index=False)

# lee los archivos 
# Lee el archivo de Excel 'BD Suscripción empresarial PyF.xlsx' y carga la hoja 'BD Matriculas' en el DataFrame eCesde 

eCesde = pd.read_excel('docs\BD Suscripción empresarial PyF.xlsx', sheet_name='BD Matriculas')
eCesde1 = pd.read_excel('docs\BD Suscripción empresarial PyF.xlsx', sheet_name='BD PLATZI')
data2 = pd.read_excel('Plantilla.xlsx')
data2.drop(index=data2.index.to_list(), inplace=True)
data3 = pd.read_excel('reporteFormateado.xlsx')
data4 = pd.read_excel('datosConsolidados.xlsx', sheet_name='Consolidado' , header=5)
data5 = pd.read_excel('datosConsolidados.xlsx', sheet_name='Usuarios Activos')
df = pd.read_excel('docs\Suscripción empresarial PyF - ECesde&Platzi.xlsx')
data7= pd.read_excel('Anexo rutas de formación v2 - TI .xlsx', sheet_name='BD Matriculas PyF')
data8= pd.read_excel('Anexo rutas de formación v2 - TI .xlsx', sheet_name='BD Matrículas Empresas')
empresarialPlatzi = pd.read_excel('docs\BD Suscripción empresarial Platzi.xlsx', sheet_name='BD Principal')
empresarialPlatzi1 = pd.read_excel('docs\BD Suscripción empresarial Platzi.xlsx', sheet_name='BD PLATZI')


#Limpia el documento copia
eCesde = eCesde.drop(index=eCesde.index[0:])
eCesde = eCesde.rename(columns={'Correo electrónico del asistente': 'Email'})
#Trae datos actualizados Anexo rutas de formación v2 - TI a datos
eCesde['Número de documento del asistente'] = data7['Numero Documento']  
eCesde['Número de documento del asistente'] = eCesde['Número de documento del asistente'].apply(lambda x: str(x).strip())
eCesde['Email'] = data7['Correo Electronico']
eCesde['Email'] = eCesde['Email'].str.lower()
eCesde['Nombre Completo'] = data7['Nombre']
eCesde['Número de contacto del asistente (Fijo o Celular)'] = data7['Telefono Celular']
eCesde['Tipo de documento del asistente'] = data7['Tipo Documento']
eCesde['Fecha Matricula'] =data7['FechaRegistro']
eCesde['Carpeta de evidencia Ecesde'] = data7['Ubicación Evidencia Ecesde'] 
eCesde['Carpeta de evidencia Platzi'] = data7['Ubicación Evidencia Platzi'] 
eCesde['FECHA DE CONCILIACIÓN'] = data7['Fecha Conciliación']
eCesde['Evidencias ECESDE Activación'] =data7['Evidencia Activación Ecesde']
eCesde['Evidencias ECESDE Progreso'] =data7['Evidencia Avance Ecesde']
eCesde['Evidencias PLATZI Activación'] =data7['Evidencia Activación Platzi']
eCesde['Evidencias PLATZI Progreso'] =data7['Evidencia Avance Platzi']



#Escribe los datos obtenidos en el Archivo de 'datos.xlsx' en la hoja 'BD Matriculas'
with pd.ExcelWriter('docs\BD Suscripción empresarial PyF.xlsx', if_sheet_exists='replace',mode='a') as writer:
    eCesde.to_excel(writer, sheet_name='BD Matriculas', index=False)



data2['Documento'] = eCesde['Número de documento del asistente']
data2['Email'] = eCesde['Email']
seriesFecha = []

# # Bucle para iterar sobre cada documento en data2['Documento']
for document in data2['Documento']:
    document = int(document)# Convierte el valor de 'document' a entero

    # Verifica si el valor de 'document' es NaN
    if(math.isnan(document)):# Convierte nuevamente el valor de 'document' a entero
        print('')
        # Establece 'Activo Ecesde' como 'SI' para el documento actual
    else:
        document = int(document)
    fila = data3.loc[data3['idnumber'] == str(document)]['days'].values
    if(len(fila) != 0):
        if(fila[0] != 'NO' and fila[0] != 'N/A'):
            data2.loc[data2['Documento'] == document, 'Activo Ecesde'] = 'SI'
        else:
            data2.loc[data2['Documento'] == document, 'Activo Ecesde'] = 'NO'

data4['Titulo del curso'] = data4['Titulo del curso'].fillna('N/A')# Rellena los valores faltantes en la columna 'Titulo del curso' de data4 con 'N/A'

# Elimina las filas de data4 donde la columna
data4.drop(data4[data4['Fecha de aprobación'] < date(day=28, month=9, year=2022)].index, inplace = True)
data4.drop(data4[data4['Último progreso del estudiante'] < date(day=28, month=9, year=2022)].index, inplace = True)
data4.drop(data4[data4['Id del curso'] == 'El estudiante no presenta actividad'].index, inplace = True)
data4.drop(data4[data4['Titulo del curso'].str.lower().str.contains('mintic')].index, inplace = True)
data4.drop(data4[data4['Titulo del curso'].str.lower().str.contains('appsco')].index, inplace = True)
data4.drop(data4[data4['Titulo del curso'].str.lower().str.contains('voz a voz')].index, inplace = True)
data4.drop(data4[data4['Titulo del curso'].str.lower().str.contains('beca')].index, inplace = True)
data4.drop(data4[data4['Titulo del curso'].str.lower().str.contains('maestros')].index, inplace = True)
data4.drop(data4[data4['Titulo del curso'].str.lower().str.contains('curso de introducción a platzi')].index, inplace = True)

data4['Email'] = data4['Email'].str.strip()# Elimina los espacios en blanco al principio y al final de los valores en la columna 'Email' de data4

data4['Último progreso del estudiante'] = pd.to_datetime(data4['Último progreso del estudiante']) # Convierte la columna 'Último progreso del estudiante' de data4 en formato de fecha y hora
data4['Último progreso del estudiante'] = data4['Último progreso del estudiante'].dt.strftime('%d/%m/%Y')
# data4['Progreso del curso (%)'] = data4['Progreso del curso (%)'].apply(lambda x: float(x.replace('%','')))

with pd.ExcelWriter('docs\BD Suscripción empresarial PyF.xlsx', if_sheet_exists='replace', mode='a') as writer:
     data4.to_excel(writer, sheet_name='BD PLATZI', index=False)  # Guarda los datos del DataFrame data4 en la hoja 'BD PLATZI' del archivo de Excel 'BD Suscripción empresarial PyF.xlsx'

with pd.ExcelWriter('docs\BD Suscripción empresarial Platzi.xlsx', if_sheet_exists='replace', mode='a') as writer:
     data4.to_excel(writer, sheet_name='BD PLATZI', index=False)



data4['Progreso del curso (%)'] = data4['Progreso del curso (%)'].fillna('N/A')
data4.drop(data4[data4['Progreso del curso (%)'] < 5].index, inplace= True)
data4.drop(data4[data4['Progreso del curso (%)'] == 'N/A'].index, inplace= True)
data4 = data4.drop_duplicates(subset=['Email'])

for correo in data4['Email']:
     correo = str(correo).lower()
     data2.loc[data2['Email'].str.lower() == correo, 'Progreso Platzi'] = 'SI'
data2['Progreso Platzi'] = data2['Progreso Platzi'].fillna('NO')

data4['Progreso del curso (%)'] = data4['Progreso del curso (%)'].apply(lambda x: '{:.0}'.format(x/100))


data5['Activo Platzi'] = 'SI'

data5.insert(2, 'Activo Platzi', data5.pop('Activo Platzi'))

for correo in data5['Email']:
     correo = str(correo).lower()
     data2.loc[data2['Email'].str.lower() == correo, 'Activo Platzi'] = 'SI'
data2['Activo Platzi'] = data2['Activo Platzi'].fillna('NO')


data5['Email'] = data5['Email'].apply(lambda x: x.strip())
data5['Fecha de envio de invitación'] = data5['Fecha de envio de invitación'].dt.strftime('%d/%m/%Y')
data5['Fecha de Activación'] = data5['Fecha de Activación'].dt.strftime('%d/%m/%Y')
data5['Ultima fecha de ingreso'] = data5['Ultima fecha de ingreso'].dt.strftime('%d/%m/%Y')

with pd.ExcelWriter('datosConsolidados.xlsx', if_sheet_exists='replace', mode='a') as writer:
     data4.to_excel(writer, sheet_name='Consolidado', index=False)
     data5.to_excel(writer, sheet_name='Usuarios activos', index=False)

for cedula in data7['Numero Documento']:
    aprobado = data7.loc[data7['Numero Documento'] == cedula]['Aprobado comfama'].values
    data2.loc[data2['Documento'] == cedula, 'Aprobados Comfama'] = aprobado
data2['Aprobados Comfama'] = data2['Aprobados Comfama'].fillna('NO')

with pd.ExcelWriter('Plantilla.xlsx') as writer:
    data2.to_excel(writer, index=False)

for cedula in data7['Numero Documento']:
    evidenciaECESDE  = data7.loc[data7['Numero Documento'] == cedula]['Ubicación Evidencia Ecesde'].values
    eCesde.loc[eCesde['Número de documento del asistente'] == cedula, 'Carpeta de evidencia Ecesde'] = evidenciaECESDE
                   
# # # # df.pop('Nombre')
# # # # df.pop('E-mail')
# # # # df.pop('Fecha de inscripción')
# # # # df.pop('Actividades Completadas')
# # # # df.pop('Actividades Asignadas')
# # # # df.pop('#')
df['username']=df['username'].apply(str)
eCesde['Número de documento del asistente']=eCesde['Número de documento del asistente'].apply(str)
for cedula in df['username']:
    porcentaje = df.loc[df['username'] == cedula]['course_completed'].values
    eCesde.loc[eCesde['Número de documento del asistente'] == cedula, 'E-mail Marketing y Fundamentos de CRM Progreso'] = porcentaje
print(eCesde['E-mail Marketing y Fundamentos de CRM Progreso'])
eCesde['Fecha Matricula'] = pd.to_datetime(eCesde['Fecha Matricula'], errors='coerce')
eCesde['Fecha Matricula'] = eCesde['Fecha Matricula'].dt.strftime('%d/%m/%Y')
eCesde['E-mail Marketing y Fundamentos de CRM Progreso'] = eCesde['E-mail Marketing y Fundamentos de CRM Progreso'].str.replace('%', '').apply(lambda x: float(x)/100)
eCesde['E-mail Marketing y Fundamentos de CRM Progreso'] = eCesde['E-mail Marketing y Fundamentos de CRM Progreso'].fillna(0)
eCesde.loc[eCesde['E-mail Marketing y Fundamentos de CRM Progreso'] == 1, 'E-mail Marketing y Fundamentos de CRM Certificado'] = 'SI'
eCesde['E-mail Marketing y Fundamentos de CRM Certificado'] = eCesde['E-mail Marketing y Fundamentos de CRM Certificado'].fillna('NO')
empresarialPlatzi.loc[empresarialPlatzi['Propósito de vida Progreso%'] == 1, 'Propósito de vida certificado'] = 'SI'

eCesde.loc[eCesde['E-mail Marketing y Fundamentos de CRM Progreso'] > 0, 'Cursos en progreso por estudiante'] = 1
eCesde.loc[eCesde['E-mail Marketing y Fundamentos de CRM Progreso'] == 0, 'Cursos en progreso por estudiante'] = 0


eCesde.loc[eCesde['E-mail Marketing y Fundamentos de CRM Certificado'] == 'SI', 'Certificados por estudiante'] = 1
eCesde.loc[eCesde['E-mail Marketing y Fundamentos de CRM Certificado'] == 'NO', 'Certificados por estudiante'] = 0

# Iterar sobre cada documento en data2 y actualizar los valores correspondientes en eCesde
for doc in data2['Documento']:
    celda = data2.loc[data2['Documento'] == doc]['Activo Ecesde'].values
    eCesde.loc[eCesde['Número de documento del asistente'] == doc, 'Activo ECESDE'] = celda
    celda = data2.loc[data2['Documento'] == doc]['Activo Platzi'].values
    eCesde.loc[eCesde['Número de documento del asistente'] == doc, 'Activo PLATZI'] = celda
    celda = data2.loc[data2['Documento'] == doc]['Progreso Platzi'].values
    eCesde.loc[eCesde['Número de documento del asistente'] == doc, 'Progreso apto para facturar PLATZI'] = celda
    celda = data2.loc[data2['Documento'] == doc]['Aprobados Comfama'].values
    eCesde.loc[eCesde['Número de documento del asistente'] == doc, 'Aprobado por COMFAMA'] = celda

eCesde.loc[eCesde['E-mail Marketing y Fundamentos de CRM Progreso'] > 0.00, 'Progreso apto para facturar ECESDE'] = 'SI'
eCesde['Progreso apto para facturar ECESDE'] = eCesde['Progreso apto para facturar ECESDE'].fillna('NO')

eCesde.loc[(eCesde['Activo PLATZI'] == 'SI') & (eCesde['Activo ECESDE'] == 'SI'), 'Activo en ambas plataformas'] = 'SI'
eCesde['Activo en ambas plataformas'] = eCesde['Activo en ambas plataformas'].fillna('NO')

eCesde.loc[(eCesde['Progreso apto para facturar ECESDE'] == 'SI') | (eCesde['Progreso apto para facturar PLATZI'] == 'SI'), 'Progreso en alguna de las 2 plataformas'] = 'SI'
eCesde['Progreso en alguna de las 2 plataformas'] = eCesde['Progreso en alguna de las 2 plataformas'].fillna('NO')

eCesde.loc[(eCesde['Activo en ambas plataformas'] == 'SI') & (eCesde['Progreso en alguna de las 2 plataformas'] == 'SI') & (eCesde['Aprobado por COMFAMA'] == 'SI'), 'CONCILIACIÓN'] = 'CONCILIAR'
eCesde['CONCILIACIÓN'] = eCesde['CONCILIACIÓN'].fillna('NO CONCILIAR')
eCesde = eCesde.rename(columns={'Email': 'Correo electrónico del asistente'})

# Creamos tres DataFrames vacíos para almacenar los datos filtrados por estado del curso
tablaAprobado = pd.DataFrame()
tablaProgreso = pd.DataFrame()
tablaCompletado = pd.DataFrame()


# Filtramos los datos para obtener los registros con estado 'Aprobado'
tablaAprobado['Nombre'] = eCesde1.loc[eCesde1['Estado del curso'] == 'Aprobado', 'Nombre completo']
tablaAprobado['Estado'] = eCesde1.loc[eCesde1['Estado del curso'] == 'Aprobado', 'Estado del curso']

tablaProgreso['Nombre'] = eCesde1.loc[eCesde1['Estado del curso'] == 'En progreso', 'Nombre completo']
tablaProgreso['Estado'] = eCesde1.loc[eCesde1['Estado del curso'] == 'En progreso', 'Estado del curso']

tablaCompletado['Nombre'] = eCesde1.loc[eCesde1['Estado del curso'] == 'Completado', 'Nombre completo']
tablaCompletado['Estado'] = eCesde1.loc[eCesde1['Estado del curso'] == 'Completado', 'Estado del curso']

# Eliminamos los duplicados en cada tabla según el nombre de la persona
tablaProgreso.drop_duplicates(subset='Nombre', inplace=True)
tablaAprobado.drop_duplicates(subset='Nombre', inplace=True)
tablaCompletado.drop_duplicates(subset='Nombre', inplace=True)

# Agrupamos los registros por estado y contamos la cantidad de registros en cada grupo
tablaAprobado = tablaAprobado.groupby('Estado')['Estado'].count().reset_index(name='Conteo')
tablaProgreso = tablaProgreso.groupby('Estado')['Estado'].count().reset_index(name='Conteo')
tablaCompletado = tablaCompletado.groupby('Estado')['Estado'].count().reset_index(name='Conteo')

# Crear un DataFrame vacío llamado "tabla"
tabla = pd.DataFrame()

# Concatenar los DataFrames "tablaAprobado", "tablaCompletado" y "tablaProgreso" en el DataFrame "tabla"
tabla = pd.concat([tablaAprobado, tablaCompletado, tablaProgreso])

# Crear un DataFrame "tablaT" con el conteo de registros agrupados por el estado del curso en el DataFrame "eCesde1"
tablaT = eCesde1.groupby('Estado del curso')['Estado del curso'].count().reset_index(name='conteoTotal')

tablaT.rename(columns={'Estado del curso': 'Estado'}, inplace=True)

tabla = pd.merge(tabla, tablaT, on='Estado')

# Guardar el DataFrame "tabla" en un archivo de Excel llamado 'BD Suscripción empresarial PyF.xlsx', en la hoja 'Hoja1'
with pd.ExcelWriter('docs\BD Suscripción empresarial PyF.xlsx', if_sheet_exists='replace',mode='a') as writer:
    tabla.to_excel(writer, sheet_name='Hoja1', index=False)


with pd.ExcelWriter('docs\BD Suscripción empresarial PyF.xlsx', if_sheet_exists='replace',mode='a') as writer:
    eCesde.to_excel(writer, sheet_name='BD Matriculas', index=False)



empresarialPlatzi = empresarialPlatzi.rename(columns={'Documento': 'idnumber'})

empresarialPlatzi = empresarialPlatzi.drop(index=empresarialPlatzi.index[0:])

empresarialPlatzi = empresarialPlatzi.rename(columns={'Documento': 'username'})
# Trae datos actualizados Anexo rutas de formación v2 - TI a datos
empresarialPlatzi['Tipo de documento del asistente'] = data8['Tipo de documento']
empresarialPlatzi['username'] = data8['Número de documento']
empresarialPlatzi['username']= empresarialPlatzi['username'].apply(lambda x: str(x)[:-2])
empresarialPlatzi['Primer nombre del asistente'] = data8['NOMBRE1']
empresarialPlatzi['Segundo nombre del asistente'] = data8['NOMBRE2']
empresarialPlatzi['Primer Apellido del asistente'] = data8['APELLIDO1']
empresarialPlatzi['Segundo Apellido del asistente'] = data8['APELLIDO2']
empresarialPlatzi['Email'] = data8['Correo electrónico']
empresarialPlatzi['Email'] = empresarialPlatzi['Email'].str.lower()
empresarialPlatzi['Número de contacto del asistente (Fijo o Celular)'] = data8['Teléfono celular']
empresarialPlatzi['Fecha Matricula'] = data8.iloc[:,24]
empresarialPlatzi['Fecha Matricula'] = empresarialPlatzi['Fecha Matricula'].replace({'20212': '2021'}, regex=True)
empresarialPlatzi['Fecha Matricula'] = pd.to_datetime(empresarialPlatzi['Fecha Matricula'])
empresarialPlatzi['Fecha Matricula'] = empresarialPlatzi['Fecha Matricula'].apply(lambda x: x.strftime('%d/%m/%Y'))
empresarialPlatzi['FECHA DE CONCILIACIÓN'] =data8['Fecha Conciliación']
empresarialPlatzi['Evidencias ECESDE Activación'] =data8['Evidencia Activación Ecesde']
empresarialPlatzi['Evidencias ECESDE Progreso'] =data8['Evidencia Avance Ecesde']
empresarialPlatzi['Carpeta de evidencia Ecesde'] =data8['Ubicación Evidencia Ecesde'] 
empresarialPlatzi['Evidencias PLATZI Activación'] =data8['Evidencia Activación Platzi']
empresarialPlatzi['Evidencias PLATZI Progreso'] =data8['Evidencia Avance Platzi']
empresarialPlatzi['Carpeta de evidencia Platzi'] =data8['Ubicación Evidencia Platzi'] 



# Lectura de los archivos Excel para reportes
reporteLiderandote = pd.read_excel('docs\curso2.xlsx')
reporteProposito = pd.read_excel('docs\curso16.xlsx')
reporteEmociones =  pd.read_excel('docs\curso9.xlsx')
reporteInnovacion =  pd.read_excel('docs\curso6.xlsx')
reporteCambio =  pd.read_excel('docs\curso10.xlsx')
reporteSer =  pd.read_excel('docs\curso1.xlsx')
reporteNeuro =  pd.read_excel('docs\curso5.xlsx')
reporteCoaching =  pd.read_excel('docs\curso11.xlsx')
reporteOrtografia =  pd.read_excel('docs\curso14.xlsx')
reporteExcel =  pd.read_excel('docs\curso7.xlsx')
reporteMacros =  pd.read_excel('docs\curso4.xlsx')
reporteBI =  pd.read_excel('docs\curso13.xlsx')
reporteServicio =  pd.read_excel('docs\curso8.xlsx')
reporteAdministracion =  pd.read_excel('docs\curso3.xlsx')
reporteVBA =  pd.read_excel('docs\curso12.xlsx')
reporteFinancieroExcel =  pd.read_excel('docs\curso15.xlsx')
reporteFinancieroExcel.drop_duplicates(subset='username',inplace=True)
with pd.ExcelWriter('docs\curso15.xlsx', if_sheet_exists='replace',mode='a') as writer:
     reporteFinancieroExcel.to_excel(writer, sheet_name='Hoja1', index=False)
# Elimina duplicados en la columna 'username' del reporteFinancieroExcel
empresarialPlatzi['username'] = empresarialPlatzi['username'].astype('str')

reporteLiderandote['username']=reporteLiderandote['username'].str.strip()
reporteLiderandote['username'] = reporteLiderandote['username'].astype('str')
merged_1 = empresarialPlatzi.merge(reporteLiderandote, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Liderándote para la vida Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Liderándote para la vida Progreso%'] = empresarialPlatzi['Liderándote para la vida Progreso%'].fillna('0%')
empresarialPlatzi['Liderándote para la vida Progreso%'] = empresarialPlatzi['Liderándote para la vida Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Liderándote para la vida Progreso%'] == 1, 'Liderándote para la vida certificado'] = 'SI'
empresarialPlatzi['Liderándote para la vida certificado'] = empresarialPlatzi['Liderándote para la vida certificado'].fillna('NO')
empresarialPlatzi['Liderándote para la vida Progreso%'] = empresarialPlatzi['Liderándote para la vida Progreso%'].fillna(0)

print('Finalizado 1')
# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Liderándote para la vida
reporteProposito['username'] = reporteProposito['username'].astype('str')
reporteProposito['username']=reporteProposito['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteProposito, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Propósito de vida Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Propósito de vida Progreso%'] = empresarialPlatzi['Propósito de vida Progreso%'].fillna('0%')
empresarialPlatzi['Propósito de vida Progreso%'] = empresarialPlatzi['Propósito de vida Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Propósito de vida Progreso%'] == 1, 'Propósito de vida certificado'] = 'SI'
empresarialPlatzi['Propósito de vida certificado'] = empresarialPlatzi['Propósito de vida certificado'].fillna('NO')
print('Finalizado 2')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Gestion de las emociones
reporteEmociones['username'] = reporteEmociones['username'].astype('str')
reporteEmociones['username']=reporteEmociones['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteEmociones, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Gestion de las emociones Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Gestion de las emociones Progreso%'] = empresarialPlatzi['Gestion de las emociones Progreso%'].fillna('0%')
empresarialPlatzi['Gestion de las emociones Progreso%'] = empresarialPlatzi['Gestion de las emociones Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Gestion de las emociones Progreso%'] == 1, 'Gestion de las emociones certificado'] = 'SI'
empresarialPlatzi['Gestion de las emociones certificado'] = empresarialPlatzi['Gestion de las emociones certificado'].fillna('NO')
print('Finalizado 3')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Creatividad e innovación
reporteInnovacion['username'] = reporteInnovacion['username'].astype('str')
reporteInnovacion['username']=reporteInnovacion['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteInnovacion, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Creatividad e innovación Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Creatividad e innovación Progreso%'] = empresarialPlatzi['Creatividad e innovación Progreso%'].fillna('0%')
empresarialPlatzi['Creatividad e innovación Progreso%'] = pd.to_numeric(empresarialPlatzi['Creatividad e innovación Progreso%'].str.replace('%', '')).apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Creatividad e innovación Progreso%'] == 1, 'Creatividad e innovación certificado'] = 'SI'
empresarialPlatzi['Creatividad e innovación certificado'] = empresarialPlatzi['Creatividad e innovación certificado'].fillna('NO')
print('Finalizado 4')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Gestión del cambio
reporteCambio['username'] = reporteCambio['username'].astype('str')
reporteCambio['username']=reporteCambio['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteCambio, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Gestión del cambio Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Gestión del cambio Progreso%'] = empresarialPlatzi['Gestión del cambio Progreso%'].fillna('0%')
empresarialPlatzi['Gestión del cambio Progreso%'] = empresarialPlatzi['Gestión del cambio Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Gestión del cambio Progreso%'] == 1, 'Gestión del cambio certificado'] = 'SI'
empresarialPlatzi['Gestión del cambio certificado'] = empresarialPlatzi['Gestión del cambio certificado'].fillna('NO')
print('Finalizado 5')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Competencias del ser para el desarrollo humano
reporteSer['username'] = reporteSer['username'].astype('str')
reporteSer['username'] = reporteSer['username'].astype('str')
reporteSer['username']=reporteSer['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteSer, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Competencias del ser para el desarrollo humano Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Competencias del ser para el desarrollo humano Progreso%'] = empresarialPlatzi['Competencias del ser para el desarrollo humano Progreso%'].fillna('0%')
empresarialPlatzi['Competencias del ser para el desarrollo humano Progreso%'] = empresarialPlatzi['Competencias del ser para el desarrollo humano Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Competencias del ser para el desarrollo humano Progreso%'] == 1, 'Competencias del ser para el desarrollo humano certificado'] = 'SI'
empresarialPlatzi['Competencias del ser para el desarrollo humano certificado'] = empresarialPlatzi['Competencias del ser para el desarrollo humano certificado'].fillna('NO')
print('Finalizado 6')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Programación neurolingüística
reporteNeuro['username'] = reporteNeuro['username'].astype('str')
reporteNeuro['username']=reporteNeuro['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteNeuro, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Programación neurolingüística Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Programación neurolingüística Progreso%'] = empresarialPlatzi['Programación neurolingüística Progreso%'].fillna('0%')
empresarialPlatzi['Programación neurolingüística Progreso%'] = empresarialPlatzi['Programación neurolingüística Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Programación neurolingüística Progreso%'] == 1, 'Programación neurolingüística certificado'] = 'SI'
empresarialPlatzi['Programación neurolingüística certificado'] = empresarialPlatzi['Programación neurolingüística certificado'].fillna('NO')
print('Finalizado 7')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Inteligencia emocional y coaching
reporteCoaching['username'] = reporteCoaching['username'].astype('str')
reporteCoaching['username']=reporteCoaching['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteCoaching, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Inteligencia emocional y coaching Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Inteligencia emocional y coaching Progreso%'] = empresarialPlatzi['Inteligencia emocional y coaching Progreso%'].fillna('0%')
empresarialPlatzi['Inteligencia emocional y coaching Progreso%'] = empresarialPlatzi['Inteligencia emocional y coaching Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Inteligencia emocional y coaching Progreso%'] == 1, 'Inteligencia emocional y coaching certificado'] = 'SI'
empresarialPlatzi['Inteligencia emocional y coaching certificado'] = empresarialPlatzi['Inteligencia emocional y coaching certificado'].fillna('NO')
print('Finalizado 8')


# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Redacción y ortografía
reporteOrtografia['username'] = reporteOrtografia['username'].astype('str')
reporteOrtografia['username']=reporteOrtografia['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteOrtografia, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Redacción y ortografía Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Redacción y ortografía Progreso%'] = empresarialPlatzi['Redacción y ortografía Progreso%'].fillna('0%')
empresarialPlatzi['Redacción y ortografía Progreso%'] = empresarialPlatzi['Redacción y ortografía Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Redacción y ortografía Progreso%'] == 1, 'Redacción y ortografía certificado'] = 'SI'
empresarialPlatzi['Redacción y ortografía certificado'] = empresarialPlatzi['Redacción y ortografía certificado'].fillna('NO')
print('Finalizado 9')


# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Herramientas esenciales de Excel
reporteExcel['username'] = reporteExcel['username'].astype('str')
reporteExcel['username']=reporteExcel['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteExcel, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Herramientas esenciales de Excel Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Herramientas esenciales de Excel Progreso%'] = empresarialPlatzi['Herramientas esenciales de Excel Progreso%'].fillna('0%')
empresarialPlatzi['Herramientas esenciales de Excel Progreso%'] = empresarialPlatzi['Herramientas esenciales de Excel Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Herramientas esenciales de Excel Progreso%'] == 1, 'Herramientas esenciales de Excel certificado'] = 'SI'
empresarialPlatzi['Herramientas esenciales de Excel certificado'] = empresarialPlatzi['Herramientas esenciales de Excel certificado'].fillna('NO')
print('Finalizado 10')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Automatización de información con el grabador de macros
reporteMacros['username'] = reporteMacros['username'].astype('str')
reporteMacros['username']=reporteMacros['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteMacros, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Automatización de información con el grabador de macros Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Automatización de información con el grabador de macros Progreso%'] = empresarialPlatzi['Automatización de información con el grabador de macros Progreso%'].fillna('0%')
empresarialPlatzi['Automatización de información con el grabador de macros Progreso%'] = empresarialPlatzi['Automatización de información con el grabador de macros Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Automatización de información con el grabador de macros Progreso%'] == 1, 'Automatización de información con el grabador de macros certificado'] = 'SI'
empresarialPlatzi['Automatización de información con el grabador de macros certificado'] = empresarialPlatzi['Automatización de información con el grabador de macros certificado'].fillna('NO')
print('Finalizado 11')


# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Análisis de datos con Power BI
reporteBI['username'] = reporteBI['username'].astype('str')
reporteBI['username']=reporteBI['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteBI, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Análisis de datos con Power BI Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Análisis de datos con Power BI Progreso%'] = empresarialPlatzi['Análisis de datos con Power BI Progreso%'].fillna('0%')
empresarialPlatzi['Análisis de datos con Power BI Progreso%'] = empresarialPlatzi['Análisis de datos con Power BI Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Análisis de datos con Power BI Progreso%'] == 1, 'Análisis de datos con Power BI certificado'] = 'SI'
empresarialPlatzi['Análisis de datos con Power BI certificado'] = empresarialPlatzi['Análisis de datos con Power BI certificado'].fillna('NO')
print('Finalizado 12')

# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Curso en excelencia en el servicio
reporteServicio['username'] = reporteServicio['username'].astype('str')
reporteServicio['username'] = reporteServicio['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteServicio, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Curso en excelencia en el servicio Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Curso en excelencia en el servicio Progreso%'] = empresarialPlatzi['Curso en excelencia en el servicio Progreso%'].fillna('0%')
empresarialPlatzi['Curso en excelencia en el servicio Progreso%'] = empresarialPlatzi['Curso en excelencia en el servicio Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Curso en excelencia en el servicio Progreso%'] == 1, 'Curso en excelencia en el servicio certificado'] = 'SI'
empresarialPlatzi['Curso en excelencia en el servicio certificado'] = empresarialPlatzi['Curso en excelencia en el servicio certificado'].fillna('NO')
print('Finalizado 13')


# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Administración desde cero
reporteAdministracion['username'] = reporteAdministracion['username'].astype('str')
reporteAdministracion['username'] = reporteAdministracion['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteAdministracion, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Administración desde cero Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Administración desde cero Progreso%'] = empresarialPlatzi['Administración desde cero Progreso%'].fillna('0%')
empresarialPlatzi['Administración desde cero Progreso%'] = empresarialPlatzi['Administración desde cero Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Administración desde cero Progreso%'] == 1, 'Administración desde cero certificado'] = 'SI'
empresarialPlatzi['Administración desde cero certificado'] = empresarialPlatzi['Administración desde cero certificado'].fillna('NO')
print('Finalizado 14')


# Realiza manipulaciones en el DataFrame empresarialPlatzi relacionadas con el reporte de Administración desde cero
reporteVBA['username'] = reporteVBA['username'].astype('str')
reporteVBA['username']=reporteVBA['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteVBA, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Macros en Excel programando con VBA Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Macros en Excel programando con VBA Progreso%'] = empresarialPlatzi['Macros en Excel programando con VBA Progreso%'].fillna('0%')
empresarialPlatzi['Macros en Excel programando con VBA Progreso%'] = empresarialPlatzi['Macros en Excel programando con VBA Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Macros en Excel programando con VBA Progreso%'] == 1, 'Macros en Excel programando con VBA certificado'] = 'SI'
empresarialPlatzi['Macros en Excel programando con VBA certificado'] = empresarialPlatzi['Macros en Excel programando con VBA certificado'].fillna('NO')
print('Finalizado 15')

reporteFinancieroExcel['username'] = reporteFinancieroExcel['username'].astype('str')
reporteFinancieroExcel['username']=reporteFinancieroExcel['username'].str.strip()
merged_1 = empresarialPlatzi.merge(reporteFinancieroExcel, on='username', how='left')
empresarialPlatzi.loc[merged_1['course_completed'].notna(), 'Funciones Financieras en Excel Progreso%'] = merged_1['course_completed']
empresarialPlatzi['Funciones Financieras en Excel Progreso%'] = empresarialPlatzi['Funciones Financieras en Excel Progreso%'].fillna('0%')
empresarialPlatzi['Funciones Financieras en Excel Progreso%'] = empresarialPlatzi['Funciones Financieras en Excel Progreso%'].str.replace('%', '').apply(lambda x: float(x)/100)
empresarialPlatzi.loc[empresarialPlatzi['Funciones Financieras en Excel Progreso%'] == 1, 'Funciones Financieras en Excel certificado'] = 'SI'
empresarialPlatzi['Funciones Financieras en Excel certificado'] = empresarialPlatzi['Funciones Financieras en Excel certificado'].fillna('NO')
print('Finalizado 16')

data4 = pd.read_excel('datosConsolidados.xlsx', sheet_name='Consolidado')

# Columna Activo ECESDE
for document in empresarialPlatzi['username']:
    fila = data3.loc[data3['idnumber'] == str(document)]['days'].values
    if(len(fila) != 0):
        if(fila[0] != 'NO' and fila[0] != 'N/A'):
            empresarialPlatzi.loc[empresarialPlatzi['username'] == document, 'Activo ECESDE'] = 'SI'
        else:
            empresarialPlatzi.loc[empresarialPlatzi['username'] == document, 'Activo ECESDE'] = 'NO'
empresarialPlatzi['Activo ECESDE'] = empresarialPlatzi['Activo ECESDE'].fillna('NO')


# Columna Activo PLATZI
for correo in data5['Email']:
     correo = str(correo).lower()
     empresarialPlatzi.loc[empresarialPlatzi['Email'] == correo, 'Activo PLATZI'] = 'SI'
empresarialPlatzi['Activo PLATZI'] = empresarialPlatzi['Activo PLATZI'].fillna('NO')

# Columna Progreso apto para facturar ECESDE
empresarialPlatzi.loc[(empresarialPlatzi['Liderándote para la vida Progreso%'] >= 0.05) | (empresarialPlatzi['Propósito de vida Progreso%'] >= 0.05) | (empresarialPlatzi['Gestion de las emociones Progreso%'] >= 0.05)
| (empresarialPlatzi['Creatividad e innovación Progreso%'] >= 0.05)| (empresarialPlatzi['Gestión del cambio Progreso%'] >= 0.05) | (empresarialPlatzi['Competencias del ser para el desarrollo humano Progreso%'] >= 0.05)
| (empresarialPlatzi['Programación neurolingüística Progreso%'] >= 0.05) | (empresarialPlatzi['Inteligencia emocional y coaching Progreso%'] >= 0.05) | (empresarialPlatzi['Redacción y ortografía Progreso%'] >= 0.05)
| (empresarialPlatzi['Herramientas esenciales de Excel Progreso%'] >= 0.05) | (empresarialPlatzi['Automatización de información con el grabador de macros Progreso%'] >= 0.05) | (empresarialPlatzi['Análisis de datos con Power BI Progreso%'] >= 0.05)
| (empresarialPlatzi['Curso en excelencia en el servicio Progreso%'] >= 0.05) | (empresarialPlatzi['Administración desde cero Progreso%'] >= 0.05) | (empresarialPlatzi['Macros en Excel programando con VBA Progreso%'] >= 0.05) 
| (empresarialPlatzi['Funciones Financieras en Excel Progreso%'] >= 0.05) ,'Progreso apto para facturar ECESDE'] = 'SI'
empresarialPlatzi['Progreso apto para facturar ECESDE'] = empresarialPlatzi['Progreso apto para facturar ECESDE'].fillna('NO')

# Columna Progreso apto para facturar PLATZI
for correo in data4['Email']:
     correo = str(correo).lower()
     empresarialPlatzi.loc[empresarialPlatzi['Email'].str.lower() == correo, 'Progreso apto para facturar PLATZI'] = 'SI'
empresarialPlatzi['Progreso apto para facturar PLATZI'] = empresarialPlatzi['Progreso apto para facturar PLATZI'].fillna('NO')

# Columna Activo en ambas plataformas
empresarialPlatzi.loc[(empresarialPlatzi['Activo ECESDE'] == 'SI') & (empresarialPlatzi['Activo PLATZI'] == 'SI'), 'Activo en ambas plataformas'] = 'SI'
empresarialPlatzi['Activo en ambas plataformas'] = empresarialPlatzi['Activo en ambas plataformas'].fillna('NO')


# Columna Progreso en alguna de las 2 plataformas
empresarialPlatzi.loc[(empresarialPlatzi['Progreso apto para facturar ECESDE'] == 'SI') | (empresarialPlatzi['Progreso apto para facturar PLATZI'] == 'SI'), 'Progreso en alguna de las 2 plataformas'] = 'SI'
empresarialPlatzi['Progreso en alguna de las 2 plataformas'] = empresarialPlatzi['Progreso en alguna de las 2 plataformas'].fillna('NO')

data8['Número de documento'] = data8['Número de documento'].astype('str')

empresarialPlatzi['Aprobado por COMFAMA'] = data8['Aprobado comfama']
empresarialPlatzi['Aprobado por COMFAMA'] = empresarialPlatzi['Aprobado por COMFAMA'].fillna('NO')
empresarialPlatzi['Aprobado por COMFAMA'] = empresarialPlatzi['Aprobado por COMFAMA'].str.upper()

# Columna CONCILIACIÓN 
empresarialPlatzi.loc[(empresarialPlatzi['Activo en ambas plataformas'] == 'SI') & (empresarialPlatzi['Progreso en alguna de las 2 plataformas'] == 'SI') & (empresarialPlatzi['Aprobado por COMFAMA'] == 'SI'), 'CONCILIACIÓN'] = 'CONCILIAR'
empresarialPlatzi['CONCILIACIÓN'] = empresarialPlatzi['CONCILIACIÓN'].fillna('NO CONCILIAR')

empresarialPlatzi = empresarialPlatzi.rename(columns={'username': 'Documento'}) # Renombra la columna 'username' a 'Documento' en el DataFrame empresarialPlatzi

with pd.ExcelWriter('docs\BD Suscripción empresarial Platzi.xlsx', if_sheet_exists='replace',mode='a') as writer:
     empresarialPlatzi.to_excel(writer, sheet_name='BD Principal', index=False)# Guarda los datos del DataFrame empresarialPlatzi en la hoja 'BD Principal' del archivo de Excel 'BD Suscripción empresarial Platzi.xlsx'

tablaAprobado1 = pd.DataFrame()
tablaProgreso1 = pd.DataFrame()
tablaCompletado1 = pd.DataFrame()

# Crea DataFrames tablaAprobado1, tablaProgreso1 y tablaCompletado1 con las columnas 'Nombre' y 'Estado' filtradas del DataFrame empresarialPlatzi1
tablaAprobado1['Nombre'] = empresarialPlatzi1.loc[empresarialPlatzi1['Estado del curso'] == 'Aprobado', 'Nombre completo']
tablaAprobado1['Estado'] = empresarialPlatzi1.loc[empresarialPlatzi1['Estado del curso'] == 'Aprobado', 'Estado del curso']

tablaProgreso1['Nombre'] = empresarialPlatzi1.loc[empresarialPlatzi1['Estado del curso'] == 'En progreso', 'Nombre completo']
tablaProgreso1['Estado'] = empresarialPlatzi1.loc[empresarialPlatzi1['Estado del curso'] == 'En progreso', 'Estado del curso']

tablaCompletado1['Nombre'] = empresarialPlatzi1.loc[empresarialPlatzi1['Estado del curso'] == 'Completado', 'Nombre completo']
tablaCompletado1['Estado'] = empresarialPlatzi1.loc[empresarialPlatzi1['Estado del curso'] == 'Completado', 'Estado del curso']

tablaProgreso1.drop_duplicates(subset='Nombre', inplace=True)
tablaAprobado1.drop_duplicates(subset='Nombre', inplace=True)
tablaCompletado1.drop_duplicates(subset='Nombre', inplace=True)
 
 # Realiza el conteo de la columna 'Estado' en los DataFrames tablaAprobado1, tablaProgreso1 y tablaCompletado1 y los agrupa por 'Estado'
tablaAprobado1 = tablaAprobado1.groupby('Estado')['Estado'].count().reset_index(name='Conteo')
tablaProgreso1 = tablaProgreso1.groupby('Estado')['Estado'].count().reset_index(name='Conteo')
tablaCompletado1 = tablaCompletado1.groupby('Estado')['Estado'].count().reset_index(name='Conteo')

tabla1 = pd.DataFrame()  # Creación de un DataFrame vacío llamado 'tabla1'

tabla1 = pd.concat([tablaAprobado1, tablaCompletado1, tablaProgreso1])

tablaT1 = empresarialPlatzi1.groupby('Estado del curso')['Estado del curso'].count().reset_index(name='conteoTotal')

tablaT1.rename(columns={'Estado del curso': 'Estado'}, inplace=True)

tabla1 = pd.merge(tabla1, tablaT1, on='Estado')

# Escritura del DataFrame 'tabla1' en un archivo Excel 'BD Suscripción empresarial Platzi.xlsx'
# en la hoja de cálculo 'Calculos', sobrescribiendo si ya existe la hoja de cálculo
with pd.ExcelWriter('docs\BD Suscripción empresarial Platzi.xlsx', if_sheet_exists='replace',mode='a') as writer:
    tabla1.to_excel(writer, sheet_name='Calculos', index=False)

print('Proceso Completado')# Impresión de mensaje de finalización del proceso






