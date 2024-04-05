# %%
import os
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta   
import getpass

# %%
nombre_usuario = getpass.getuser()
#archivo_evolutivo = pd.read_excel(f"C:/Users/{nombre_usuario}/Grupo Derco/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Instock Semanal/Evolutivo.xlsx" )
# archivo_evolutivo['Venta UMB'] = archivo_evolutivo['Venta UMB']/6
hoy = datetime.datetime.today().strftime("%Y-%m-%d")

# %%
def obtener_nombre_mes(numero_mes):
    meses = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    }
    return meses.get(numero_mes, "Mes no válido")
nombre_mes = obtener_nombre_mes(datetime.date.today().month)


# %%
dtype ={'ue': str}
valores_permitidos = ['Accesorios-Car Care','Rep.Alter.Maquinaria','Repuesto Alternativo','Baterías','Neumáticos','Lubricantes']
mara_ue = pd.read_excel(f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Instock Semanal/Maestra Aftermarket Actualizable.xlsx", usecols=['ue','Nombre Sector'], dtype= dtype)
mara_ue = mara_ue.drop_duplicates(subset=['ue'])
mara_filtrada = mara_ue[mara_ue['Nombre Sector'].isin(valores_permitidos)]
# mara_filtrada['UE'] = mara_filtrada['UE'].astype('str')

mara_filtrada.loc[:, 'ue'] = mara_filtrada['ue'].astype(str)
mara_filtrada.loc[:, 'ue'] = mara_filtrada['ue']

# %%
carpeta = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal"
carpeta_ventas = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal"
#carpeta_destino = f"C:/Users/{nombre_usuario}/Grupo Derco/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Instock Semanal/SEM {}".format(datetime.date.today().isocalendar()[1])
numero_semana = datetime.datetime.today().isocalendar()[1]

# %%
nombre_carpeta = "SEM {}".format(datetime.datetime.today().isocalendar()[1])

# Specify the path where you want to create the folder
ubi_carpeta = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificacion y Compras Procesos y Metodos/Pandas/INSTOCK"

# Combine the folder name with the path to create the full path
carp_semana_instock = os.path.join(ubi_carpeta, nombre_carpeta)

# %%
carpetadestino = os.listdir(f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal")[-2]
stock_destino = os.path.join(f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal", carpetadestino)    
#ruta = stock_destino + '/2023-10-30 - Stock.xlsx'.format(datetime.datetime.today().isocalendar()[1])
stock_destino_l = os.listdir(stock_destino)
for i  in stock_destino_l:
    if 'tock' in i and not 'iendas' in i and not 'añol' in i and 'S4' in i:
        ruta = os.path.join(stock_destino, i)
        print(ruta)

# ruta = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal/2024-01-02/2024-01-02 - Stock.XLSX"
stock = pd.read_excel(ruta,sheet_name='Sheet1', usecols=['Material', 'Libre utilización'])
# #stock = stock.iloc[stock['Almacen'] != 1600]
stock_suma = stock.groupby('Material', as_index=False).agg({'Libre utilización': 'sum'})
stock_suma['Libre utilización'].sum()

# %%
maestro = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros"
dir_maestro = os.listdir(maestro)
for c_año in dir_maestro:
    if str(datetime.datetime.today().year) in c_año:
        c_carpeta = os.path.join(maestro, c_año)
        c_mes = os.listdir(c_carpeta)
        c_arch = os.path.join(c_carpeta, c_mes[-1])
        archivos = os.listdir(c_arch)
        print(archivos)
        for a in archivos:
            if 'COD_ACTUAL_S4' in a:
                ruta_cad = os.path.join(c_arch, a)
                cadena_de_remplazo = pd.read_excel(ruta_cad, usecols= ['Nro_pieza_fabricante_1',	'Cod_Actual_1'] )
                
# ruta_cad_reemplazo = f"C:/Users/{nombre_usuario}

# %%
df = pd.merge(stock_suma, cadena_de_remplazo, left_on= 'Material', right_on='Nro_pieza_fabricante_1', how= 'left', suffixes=('Cod_Actual_1', '_Cad_reemplazo'))
df.drop('Nro_pieza_fabricante_1', axis=1, inplace=True)
df['Cod_Actual_1'].fillna(df['Material'], inplace=True)
#df['SEM'] = numero_semana
df.drop('Material', axis=1, inplace=True)
df_suma = df.groupby(['Cod_Actual_1']).agg({'Libre utilización': 'sum'}).reset_index()

# %%

fechas_ventas = []
semanas_ventas = sorted(os.listdir(carpeta), reverse=True)[1:13]
consolidado_ventas = []
columnas_ventas = ['Material', 'Venta UMB','Semana ISO'] 
for ventas in semanas_ventas:
    ruta_semana_ventas = os.path.join(carpeta_ventas, ventas)
    for ventas in os.listdir(ruta_semana_ventas):
        if "Sell" in ventas and "Out" not in ventas and 'R3' not in ventas:
            arc_ventas = os.path.join(ruta_semana_ventas, ventas)
            dfv = pd.read_excel(arc_ventas, sheet_name="Sheet1", usecols=columnas_ventas, dtype={'Material':'str'}) #.groupby('Material')['Venta UMB'].sum().reset_index()
            fechas_ventas.append(ventas[:10])


            consolidado_ventas.append(dfv)
df_final_ventas = consolidado_ventas[0]

# %%


for i in range(1, len(consolidado_ventas)):
    #df_final_ventas = pd.merge(df_final_ventas, consolidado_ventas[i], on='Material', how='left', sufixes=(f'_SEM{i}', f'_SEM{i+1}'))

    df_final_ventas = pd.concat([df_final_ventas, consolidado_ventas[i]], axis=0, ignore_index=True)

df_final_ventas = df_final_ventas.groupby(['Material', 'Semana ISO'])['Venta UMB'].sum().reset_index()
df_final_ventas.drop(['Semana ISO'], axis=1)
df_final_ventas = df_final_ventas.groupby(['Material'])['Venta UMB'].sum().reset_index()

    #df_final_ventas = df_final_ventas.applymap(lambda x: int(x) if '-' not in x else 0)
    #df_final_ventas.replace(0, np.nan, inplace=True)


    # Filtrar las filas que no contienen un guion '-'

    #agrupar por ultimo eslabon

df = pd.merge(df_final_ventas, cadena_de_remplazo, left_on= 'Material', right_on='Nro_pieza_fabricante_1', how= 'left', suffixes=('Cod_Actual_1', '_Cad_reemplazo'))
df.drop('Nro_pieza_fabricante_1', axis=1, inplace=True)
df['Cod_Actual_1'].fillna(df['Material'], inplace=True)
#df['SEM'] = numero_semana
df.drop('Material', axis=1, inplace=True)
ventas_suma = df.groupby(['Cod_Actual_1']).agg({'Venta UMB': 'sum'}).reset_index()


#nombre_archivo_ventas = "Ventas SEM {}.xlsx".format(datetime.date.today().isocalendar()[1])
#ruta_destino = os.path.join(carpeta_destino, nombre_archivo_ventas)
# df_consolidado.to_excel(ruta_destino, index=False)

# def reemplazar_negativos_con_nan(valor):
#     return valor if valor > 0 else None

# Aplicar la función solo a las columnas que lo necesitas (desde la segunda columna en adelante)
# columnas_a_modificar = df_final_ventas.columns[1:]  # Obtiene las columnas a partir de la segunda
# df_final_ventas[columnas_a_modificar] = df_final_ventas[columnas_a_modificar].applymap(reemplazar_negativos_con_nan)


# df_final_ventas['Promedio Semanal Ventas'] = df_final_ventas.iloc[:, 1:].mean(axis=1)
# columnas = df_final_ventas.columns[1:-1]
# df_final_ventas.drop(columnas, axis=1, inplace=True)


#ventas_suma.to_excel(ruta_destino, index= False)
print('Para el promedio de ventas se tomaron los archivos correspondientes a las siguientes fechas: ' + "\n" + ', '.join(str(ventas) for ventas in fechas_ventas))


# %%
ventas_suma['Venta UMB'].sum()

# %%
len(fechas_ventas)

# %%
df.shape

# %%
ventas_v_stock = pd.merge(mara_filtrada ,ventas_suma, left_on='ue',right_on='Cod_Actual_1', how= 'left', suffixes=('_Venta UMB', 'Venta UMB'))


#ventas_v_stock['SEMANA'] = datetime.datetime.today().isocalendar()[1]
#ventas_v_stock.drop('Material', axis=1, inplace=True)

#DESDE LA ANTERIOR TRAER STOCK

ventas_v_stock_2 = pd.merge(ventas_v_stock, df_suma, left_on='ue', right_on='Cod_Actual_1', how='left', suffixes=('_Libre utilización', 'Stock'))


###PRUEBA DE INCORPORAR INSTOCK EN EL CALCULO INICIAL
ventas_v_stock_2['Venta UMB'] = ventas_v_stock_2['Venta UMB'].apply(lambda x: 0 if x < 0 else x)
ventas_v_stock_2['Venta UMB'] = ventas_v_stock_2['Venta UMB']/6

ventas_v_stock_2['Libre utilización'].fillna(0, inplace=True)
ventas_v_stock_2['Venta UMB'].fillna(0, inplace=True)


def aplicar_condicion(row):
    if (row['Libre utilización'] - row['Venta UMB']) >= 0:
        return 1
    else:
        return 0

# Aplicar la función usando apply y lambda
ventas_v_stock_2['Instock'] = ventas_v_stock_2.apply(lambda row: aplicar_condicion(row), axis=1)









#ventas_v_stock.to_excel(carpeta_destino + '/BASE INSTOCK SEM {}.xlsx'.format(numero_semana,numero_semana))
ventas_v_stock_2.to_excel(f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Instock Semanal/Bases Instock/2024 Base INSTOCK {nombre_mes} sem {datetime.datetime.today().isocalendar()[1]}.xlsx")

# %%
ventas_v_stock['Venta UMB'].sum()
#rango aprox 1996323

# %%
ventas_suma

# %%
ventas_v_stock['Venta UMB'].sum()

# %%
vigencia = pd.read_excel(f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Vigencias/{hoy[:7]} Vigencia Demanda AFM.xlsx", usecols=['ue','Segmentacion Aftermarket','Categoria ClaCom','CES 01','Mayorista','Sodimac','Easy','Walmart','SMU','Tottus','Retail (AP-AG)', 'Agrupacion ClaCom','Subagrupacion ClaCom',	'Categoria ClaCom',	'Subcategoria ClaCom'])
#return cadena_de_reemplazo
#print(a)



vigencia['Segmentacion Aftermarket'] = vigencia['Segmentacion Aftermarket'].replace('SIN SEGMENTACION', 0)
vigencia['Segmentacion Aftermarket'] = vigencia['Segmentacion Aftermarket'].replace('NO APLICA', 0)


vigencia['Segmentacion Aftermarket'] = vigencia['Segmentacion Aftermarket'].fillna(0)
vigencia['Segmentacion Aftermarket'] = vigencia['Segmentacion Aftermarket'].astype(int)


def custom_max_zero(column):
    if (column == 0).all():
        return 0
    else:
        return column.max()

vigencia_agrupada = vigencia.groupby(['ue']).agg({
'Segmentacion Aftermarket': custom_max_zero,
'CES 01': 'sum',

'Mayorista': 'sum',
'Sodimac': 'sum',
'Easy': 'sum',
'Walmart': 'sum',
'SMU': 'sum',
'Tottus': 'sum',
'Retail (AP-AG)': 'sum',
'Agrupacion ClaCom': 'first',
'Subagrupacion ClaCom': 'first',
'Categoria ClaCom': 'first',
'Subcategoria ClaCom': 'first'}).reset_index()

vigencia_agrupada['Segmentacion Aftermarket'] = vigencia_agrupada['Segmentacion Aftermarket'].replace(0, 'SIN SEGMENTACION')
vigencia_agrupada.to_excel(f'C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Instock Semanal/Bases Vigencias' + f'/BASE VIGENCIAS {nombre_mes} SEM {numero_semana}.xlsx')



# %%
ventas_v_stock_2['Libre utilización'].sum()

# %%
ventas_v_stock_2[ventas_v_stock_2['Nombre Sector']=='Accesorios-Car Care']['Libre utilización'].sum()
# rango aprox = 1750786


