#!/usr/bin/env python
# coding=utf-8
import pandas as pd
import numpy as np
import os
import xlwings as xw
#import seaborn
import matplotlib.pyplot as plt
#import xlrd
#import cx_Oracle


## Determinar la longitud de la tabla a mostrar
desired_width = 320
pd.set_option('display.width', desired_width)
pd.set_option('display.max_columns', None)

## Conectar a la base de datos de Oracle
#conn_str = u'cmde_raw/raw17@172.16.120.3:1521/cnih'
#conn = cx_Oracle.connect(conn_str)
#c = conn.cursor()

#query = 'SELECT * FROM CAT_TIPO_CAMPOS'
#cat_tipo_campos = pd.read_sql(con = conn, sql=query)
#cat_tipo_campos.to_csv('CAT_TIPO_CAMPOS.csv', index = False)

#conn.close()

## Establecer el directorio de trabajo
#os.chdir('Z:\\Reportes\\37. Reservas (Anual)\\2018\\Originales')

## Funcion para encontrar los archivos .csv
#def find_filenames(path_to_dir = '.', suffix=".csv"):
#    filenames = os.listdir(path_to_dir)
#    return [filename for filename in filenames if filename.endswith(suffix)]

## Usar la funcion para encontrar los archivos que acaben en .xlsx y extraer el Anexo
#filenames = find_filenames(suffix=".xlsx")
#working_file = filter(lambda x: x.startswith('Indicadores Perfiles'), filenames)

## Importar la informacion de reservas
#col_names = pd.read_excel(working_file[0], sheet_name='PMX_Limite economico', header=4, nrows=0).columns
#types_dict = {'Region': object, 'Activo': object, 'Campo': object, 'Asignacion / Contrato': object, 'Categoria': object, 'Perfil': object}
#types_dict.update({col: float for col in col_names if col not in types_dict})
#perfiles = pd.read_excel(working_file[0], sheet_name='PMX_Limite economico', header=4, usecols = "A:F, H:BP", dtype = types_dict)
#perfiles.fillna(0, inplace=True)
#perfiles[['Region', 'Activo', 'Categoria']] = perfiles[['Region', 'Activo', 'Categoria']].astype('category')
#perfiles['id'] = perfiles.groupby(['Campo', 'Categoria']).ngroup()

## Leer los filtros de campos y categoria
os.chdir('/Users/pablo/Documents/Proyecto_Python')
col_names = pd.read_excel('Plantilla_RF.xlsm', sheet_name='Datos', header=0, nrows=0).columns
types_dict = {'Region': object, 'Activo': object, 'Campo': object, 'Asignacion / Contrato': object, 'Categoria': object, 'Perfil': object}
types_dict.update({col: float for col in col_names if col not in types_dict})

wb = xw.Book('Plantilla_RF.xlsm').sheets['Reporte']
tipo_analisis = wb.range('B6').value
campo = wb.range('B8').value
categoria = wb.range('B10').value
region_fiscal =  wb.range('B12').value
regimen = wb.range('B14').value

if tipo_analisis == 'Campo' and regimen=='Asignacion':
    ## Importar la informacion de reservas
    datos = xw.Book('Plantilla_RF.xlsm').sheets['Datos']
    perfiles = datos.range('A1').expand().options(pd.DataFrame).value.reset_index()
    mask1 = perfiles['Campo'] == campo               ## filtro por campo
    mask2 = perfiles['Categoria'] == categoria       ## filtro por categoria

    perfiles = perfiles[mask1 & mask2]
    perfiles.fillna(0, inplace=True)
    perfiles = perfiles.drop(labels = 'Total', axis=1)
    perfiles['id'] = perfiles.groupby(['Campo', 'Categoria']).ngroup()


    ## Reacomodar la tabla
    perfilesm = pd.melt(perfiles, id_vars=['Region', 'Activo', 'Campo', 'Asignacion / Contrato', 'Categoria', 'Perfil', 'id'], var_name='Año', value_name='Monto')
    perfilesm = perfilesm.groupby(['Region', 'Activo', 'Campo', 'Asignacion / Contrato', 'Categoria', 'id', 'Año', 'Perfil'])['Monto'].aggregate('sum').unstack(level = 'Perfil').reset_index()
    tipo_campo = perfilesm.groupby(['Campo', 'Categoria'])['Crudo (mb)'].aggregate('sum').reset_index()
    tipo_campo['tipo'] = tipo_campo['Crudo (mb)'].map(lambda x: 'Asociado' if x>0 else 'No Asociado')
    tipo_campo = tipo_campo[['Campo', 'Categoria', 'tipo']]
    perfilesm = perfilesm.merge(tipo_campo, how='outer', left_on = ['Campo', 'Categoria'], right_on=['Campo', 'Categoria'])
    perfilesm['Campo'] = map(lambda x: x.upper(), perfilesm['Campo'])
    perfilesm['Campo'] = perfilesm.Campo.str.normalize('NFKD').str.encode('utf-8').str.decode('ascii', 'ignore')
    perfilesm.columns = perfilesm.columns.str.replace('ó', 'o')

    ## Unir las tablas de CAT_TIPO_CAMPO con la de PERFILESM para conocer la ubicacion
    catalogo = xw.Book('Plantilla_RF.xlsm').sheets['cat_tipo_campos']
    cat_tipo_campos = catalogo.range('A1').expand().options(pd.DataFrame).value.reset_index()
    perfilesm = perfilesm.merge(cat_tipo_campos[['CAMPO','UBICACION']], how='left', left_on='Campo', right_on='CAMPO')
    perfilesm['UBICACION'] = perfilesm['UBICACION'].replace(np.nan, 'Terrestre', regex=True)  #los campos que no fueron unidos, los identificamos como Terrestres de manera manual
    perfilesm['UBICACION'] = perfilesm['UBICACION'].replace('Aguas Someras', 'Aguas someras')

    ## Asignar el año como indice en formato date
    perfilesm['Año'] = pd.to_datetime(perfilesm['Año'].astype(int), format='%Y')  #.dt.date.astype("datetime64[ns]")
    #perfilesm.set_index('Año', inplace=True)


    ## Crear DataFrame de precios de aceite, gas y condensados por año y unir al dataframe de perfilesm
    column_precios = pd.DataFrame(columns=['Año', 'precio_aceite', 'precio_gas', 'precio_condensado'])
    column_precios[['Año']] = pd.DataFrame(pd.date_range(perfilesm['Año'].min(), perfilesm['Año'].max(), freq='AS'))
    column_precios[['precio_aceite']] = float(wb.range('B17').value)
    column_precios[['precio_gas']] = float(wb.range('B19').value)
    column_precios[['precio_condensado']] = float(wb.range('B21').value)
    perfilesm = perfilesm.merge(column_precios, how='left', left_on='Año', right_on='Año')


    ## Variables generales
    tipo_cambio = float(wb.range('I6').value)
    pce_tc = float(wb.range('B24').value)
    pce_as = float(wb.range('B23').value)
    vh_as = float(wb.range('I20').value)
    vh_tc = float(wb.range('I21').value)
    vh_atg = float(wb.range('I22').value)
    vh_gna = float(wb.range('I24').value)
    vh_ap = float(wb.range('I23').value)

    iaeeh_expl = 1583.74
    iaeeh_ext = 6334.98
    area_km2 = float(wb.range('B15').value)

    valor_remanente = float(wb.range('I8').value)
    tasa_descuento = float(wb.range('I12').value)
    tasa_duc = 0.65
    tasa_isr = float(wb.range('I10').value)

    ##############################################################################################################
    ##                                              FUNCIONES                                                   ##
    ##############################################################################################################


    ## Estimar la tasa de DEXT
    def tasa_dext(tasa_aceite=perfilesm['precio_aceite'], tasa_gas=perfilesm['precio_gas'],tasa_condensado=perfilesm['precio_condensado']):
        pct_aceite = pd.Series(tasa_aceite).map(lambda x: 0.075 if x < 45.95 else ((x * 0.125) + 1.5) / 100).rename(
            'tasa_aceite')
        pct_gasoc = pd.Series(tasa_gas).map(lambda x: x / 100).rename('tasa_gasoc')
        pct_gnasoc = pd.Series(tasa_gas).map(
            lambda x: 0 if x <= 4.79 else (((x - 5) * 0.605) / x if x < 5.5 else x / 100)).rename('tasa_gnasoc')
        pct_condensado = pd.Series(tasa_condensado).map(
            lambda x: 0.05 if x < 57.44 else ((0.125 * x) - 0.025) / 100).rename('tasa_condensado')
        pct_dext = pd.concat([pct_aceite, pct_gasoc, pct_gnasoc, pct_condensado], axis=1)
        return (pct_dext)


    ## Crear funcion para estimar el Derecho de Extraccion
    def dext(tabla_base=perfilesm):
        tabla_dext = pd.concat([tabla_base, tasa_dext()], axis=1)
        tabla_dext['VCH_aceite'] = tabla_dext['Crudo (mb)'] * tabla_dext['precio_aceite'] / 1000
        tabla_dext['VCH_gas'] = tabla_dext['Gas (mmpc)'] * tabla_dext['precio_gas'] / 1000
        tabla_dext['VCH_condensado'] = tabla_dext['Condensado (mb)'] * tabla_dext['precio_condensado'] / 1000
        tabla_dext['VCH_MMUSD'] = (tabla_dext['VCH_aceite'] + tabla_dext['VCH_gas'] + tabla_dext['VCH_condensado'])
        tabla_dext['dext_aceite'] = tabla_dext['VCH_aceite'] * tabla_dext['tasa_aceite']
        tabla_dext['dext_gasoc'] = (tabla_dext.VCH_gas * tabla_dext.tasa_gasoc).where(tabla_dext.tipo == 'Asociado', other=0)
        tabla_dext['dext_gnasoc'] = (tabla_dext.VCH_gas * tabla_dext.tasa_gnasoc).where(tabla_dext.tipo == 'No Asociado', other=0)
        tabla_dext['dext_condensado'] = tabla_dext.VCH_condensado * tabla_dext.tasa_condensado
        tabla_dext['DEXTH_MMUSD'] = (tabla_dext.dext_aceite + tabla_dext.dext_gasoc + tabla_dext.dext_gnasoc + tabla_dext.dext_condensado)
        tabla_dext = tabla_dext[['Campo', 'Categoria', 'Año', 'DEXTH_MMUSD', 'VCH_aceite', 'VCH_gas', 'VCH_condensado', 'VCH_MMUSD']]
        return (tabla_dext)


    ## Crear funcion para estimar el Derecho de Exploracion
    dexpl_inicio = 1214.21
    dexpl_despues = 2903.54
    area_km2 = float(wb.range('B15').value)


    def dexpl(tabla_base=perfilesm, tabla=dext()):
        tabla_dexpl = tabla_base
        tabla_dexpl['PCEcum'] = tabla_dexpl.groupby(by=['Campo', 'Categoria'])['PCE (mb)'].cumsum()
        tabla_dexpl['etapa'] = tabla_dexpl['PCEcum'].map(lambda x: float(0) if x > 0 else float(1))
        tabla_dexpl['etapacum'] = tabla_dexpl.groupby(by=['Campo', 'Categoria'])['etapa'].cumsum()
        tabla_dexpl['etapacum'] = tabla_dexpl['etapa'] * tabla_dexpl['etapacum']
        tabla_dexpl['DEXPL_MMUSD'] = tabla_dexpl['etapacum'].map(lambda x: ((dexpl_inicio * area_km2 * 12) / tipo_cambio) / 1000000 if x < 6 and x > 0 else (((dexpl_despues * area_km2 * 12) / tipo_cambio) / 1000000 if x > 5 and x > 0 else 0))  # el cobro por derecho de exploracion cambia en el mes 61, es decir, a partir del 5to año
        tabla_dexpl = tabla_dexpl[['Campo', 'Categoria', 'Año', 'DEXPL_MMUSD']]
        return (tabla_dexpl)



    ## Crear funcion para estimar el Derecho de Utilidad Compartida

    ## Escribir la funcion que defina el limite de recuperacion de costos, con base en el VCH
    def costcap_vch(tabla):
        resultado = None
        if tabla['Activo'] == 'Aceite Terciario del Golfo':
            resultado = tabla['VCH_MMUSD'] * vh_atg
        else:
            if tabla['tipo'] =='No Asociado':
                resultado = tabla['VCH_MMUSD'] * vh_gna
            else:
                if tabla['UBICACION'] == 'Aguas someras':
                    resultado = tabla['VCH_MMUSD'] * vh_as
                elif tabla['UBICACION'] == 'Aguas profundas':
                    resultado = tabla['VCH_MMUSD'] * vh_ap
                elif tabla['UBICACION'] == 'Terrestre':
                    resultado = tabla['VCH_MMUSD'] * vh_tc
        return resultado

    ## Escribir la funcion que defina el limite de recuperacion de costos, con base en PCE
    def costcap_pce(tabla):
        resultado = None
        if tabla['tipo'] == 'No Asociado':
            resultado = 0
        else:
            if tabla['UBICACION'] == 'Aguas someras':
                resultado = (tabla['PCE (mb)'] * pce_as)/1000
            elif tabla['UBICACION'] == 'Terrestre':
                resultado = (tabla['PCE (mb)'] * pce_tc)/1000
            else:
                resultado = 0
        return resultado


    def duc(tabla_base = perfilesm, dext = dext(), dexpl = dexpl()):
        deduc_exploracion = float(wb.range('I15').value)
        deduc_desarrollo = float(wb.range('I16').value)
        deduc_infraestructura = float(wb.range('I17').value)
        periodos_deduc_desarrollo = int(1/deduc_desarrollo)  #defino los periodos a deducir la inversion de desarrollo
        periodos_deduc_infraestructura = int(1/deduc_infraestructura)  #defino los periodos a deducir la inversion de infraestructura
        tabla_duc = tabla_base.merge(dext, how='left', on = ['Campo', 'Categoria', 'Año'])
        tabla_duc = tabla_duc[['Activo','Campo', 'Categoria', 'Año', 'tipo', 'UBICACION', 'Crudo (mb)', 'Gas (mmpc)', 'Condensado (mb)', 'Costo variable (MMUSD)', 'Costos fijos (MMUSD)', 'Inversiones (MMUSD)', 'PCE (mb)', 'DEXTH_MMUSD', 'DEXPL_MMUSD', 'VCH_aceite', 'VCH_gas', 'VCH_condensado', 'VCH_MMUSD']]
        tabla_duc['Costcap_VCH'] = tabla_duc.apply(costcap_vch,axis=1)
        tabla_duc['Costcap_PCE'] = tabla_duc.apply(costcap_pce,axis=1)
        tabla_duc['Costcap_max'] = tabla_duc[['Costcap_VCH', 'Costcap_PCE']].max(axis=1)
        d = {}
        for x in range(0, periodos_deduc_desarrollo):
            d["Inversiones_{0}".format(x)] = tabla_duc.groupby(['Campo', 'Categoria'])['Inversiones (MMUSD)'].shift(x)*deduc_desarrollo
            tabla_duc['Inversiones_{0}'.format(x)] = d["Inversiones_{0}".format(x)]
        if 'Infraestructura (MMUSD)' in tabla_duc.columns.values:
            infra = {}
            for x in range(0, periodos_deduc_infraestructura):
                infra["Infraestructura_{0}".format(x)] = tabla_duc.groupby(['Campo', 'Categoria'])['Infraestructura (MMUSD)'].shift(x) * deduc_infraestructura
                tabla_duc['Infraestructura_{0}'.format(x)] = d["Infraestructura_{0}".format(x)]
            tabla_duc['Inversiones_Deduc'] = tabla_duc.iloc[:, tabla_duc.columns.str.startswith('Inversiones_')].sum(axis=1)
            tabla_duc['Infraestructura_Deduc'] = tabla_duc.iloc[:, tabla_duc.columns.str.startswith('Infraestructura_')].sum(axis=1)
            tabla_duc['GastoPorDeduc'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)'] + tabla_duc['Inversiones_Deduc'] + tabla_duc['Infraestructura_Deduc']
            tabla_duc['GastoDeducible'] = tabla_duc[['GastoPorDeduc', 'Costcap_max']].min(axis=1)
            return(tabla_duc)
        else:
            tabla_duc['Inversiones_Deduc'] = tabla_duc.iloc[:, tabla_duc.columns.str.startswith('Inversiones_')].sum(axis=1)
            tabla_duc['GastoPorDeduc'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)'] + tabla_duc['Inversiones_Deduc']
            tabla_duc['GastoDeducible'] = tabla_duc[['GastoPorDeduc', 'Costcap_max']].min(axis=1)
            return(tabla_duc)

    tabla_duc = duc()

    ## Creo una variable que divida los grupos de acuerdo al Campo y Categoria para aplicar el siguiente for loop
    tabla_duc2 = tabla_duc
    tabla_duc2 = tabla_duc2.groupby(['Campo', 'Categoria'])

    df = pd.DataFrame(columns = tabla_duc.columns)  #Creo un data.frame que tenga los mismos nombres de columnas para pegar los calculos

    ## El siguiente for-loop itera sobre cada grupo para calcular el valor acumulado a deducir cada año
    for n, g in tabla_duc2:
        print(n)
        g = g.reset_index()
        g.loc[1, 'dif_costcap_invpordeduc'] = g.loc[0, 'GastoPorDeduc'] - g.loc[0, 'GastoDeducible']    # calculo la inversion no deducida en el primer año y la aviento al segundo año
        g.loc[:, 'dif_costcap_invpordeduc'].fillna(value=0, axis=0, inplace=True)
        for i in range(1, len(g['Año'])):
            g.loc[i, 'GastoPorDeduc'] = g.loc[i, 'GastoPorDeduc'] + g.loc[i, 'dif_costcap_invpordeduc']        # sumo la inversion no deducida del primer año a la inversion por deducir del año 2
            g.loc[i+1, 'dif_costcap_invpordeduc'] = g.loc[i, 'GastoPorDeduc'] - g.loc[i, 'GastoDeducible']     # obtengo la diferencia del gasto por deducir y lo que se puede deducir del año dos
        df = df.append(g)

    df['UO_MMUSD'] = df['VCH_MMUSD'] - df['DEXTH_MMUSD'] - df['GastoDeducible']
    df['DUC_MMUSD'] = tasa_duc * df['UO_MMUSD']
    df['DERECHOS_MMUSD'] = df['DUC_MMUSD'] + df['DEXTH_MMUSD'] + df['DEXPL_MMUSD']

    tabla_duc = df.drop(labels=['GastoDeducible', 'dif_costcap_invpordeduc', 'index'], axis=1)
    tabla_duc.dropna(axis=0, how='all', inplace=True)

    ## Crear funcion para calcular el IAEEH

    def iaeeh(tabla_base = perfilesm):
        tabla_iaeeh = tabla_base
        tabla_iaeeh = tabla_base
        tabla_iaeeh['PCEcum'] = tabla_iaeeh.groupby(by=['Campo', 'Categoria'])['PCE (mb)'].cumsum()
        tabla_iaeeh['etapa'] = tabla_iaeeh['PCEcum'].map(lambda x: int(0) if x > 0 else int(1))
        tabla_iaeeh['IAEEH_MMUSD'] = tabla_iaeeh['etapa'].map(lambda x: ((iaeeh_expl*area_km2*12)/tipo_cambio)/1000000 if x==1 else ((iaeeh_ext*area_km2*12)/tipo_cambio)/1000000)
        tabla_iaeeh.loc[(tabla_iaeeh['PCE (mb)'] == 0) & (tabla_iaeeh['PCEcum'] > 0), 'IAEEH_MMUSD'] = 0   # los periodos donde ya no se esta produciendo y no son exploratorios, no se cobra el IAEEH
        tabla_iaeeh = tabla_iaeeh[['Campo', 'Categoria', 'Año', 'IAEEH_MMUSD']]
        return(tabla_iaeeh)

    ## Determinar el Impuesto sobre la Renta (ISR)
    tabla_duc = tabla_duc.merge(iaeeh(), how='left', on = ['Campo', 'Categoria', 'Año'])
    tabla_duc['Deducciones'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)'] + tabla_duc['Inversiones_Deduc']
    tabla_duc['Ingreso gravable'] = tabla_duc['VCH_MMUSD'] - tabla_duc['DERECHOS_MMUSD'] - tabla_duc['Deducciones'] - tabla_duc['IAEEH_MMUSD']


    def perdidas_acumuladas(tabla=tabla_duc):
        tabla2 = tabla.groupby(['Campo', 'Categoria'])
        df = pd.DataFrame(columns=tabla.columns)

        ## El siguiente for-loop itera sobre cada grupo para calcular el valor acumulado a deducir cada año
        for n, g in tabla2:
            print(n)
            g = g.reset_index()
            g.loc[0, 'Perdidas Acumuladas 1'] = np.min(pd.to_numeric((g.loc[0,'Ingreso gravable'], g.loc[0,'Ingreso gravable']-valor_remanente)))    # asigno si hay alguna inversion arrastrando por deducir
            for i in range(1, len(g['Año'])):
                if g.loc[i, 'Ingreso gravable'] < 0:
                    g.loc[i, 'Perdidas Acumuladas 1'] = g.loc[i-1, 'Perdidas Acumuladas 1'] + g.loc[i, 'Ingreso gravable']
                elif g.loc[i, 'Ingreso gravable'] < abs(g.loc[i-1, 'Perdidas Acumuladas 1']):
                    g.loc[i, 'Perdidas Acumuladas 1'] = g.loc[i-1, 'Perdidas Acumuladas 1'] + g.loc[i, 'Ingreso gravable']
                else:
                    g.loc[i, 'Perdidas Acumuladas 1'] = 0
            df = df.append(g)


        df['Perdidas Acumuladas 2'] = df.loc[:, 'Ingreso gravable'].rolling(min_periods=1, window=11).sum()
        df['Perdidas Acumuladas'] = df[['Perdidas Acumuladas 1', 'Perdidas Acumuladas 2']].max(axis=1)
        df['IGAP'] = np.where(df['Ingreso gravable'] <= 0, 0, np.where(df['Perdidas Acumuladas'] <= 0, 0, df[['Ingreso gravable', 'Perdidas Acumuladas']].min(axis=1)))
        return(df)

    tabla_duc = perdidas_acumuladas()

    tabla_duc['ISR_MMUSD'] = tabla_duc['IGAP'] * tasa_isr
    tabla_duc['FlujoEfectivo_MMUSD'] = tabla_duc['Ingreso gravable'] - tabla_duc['ISR_MMUSD']
    tabla_duc['FlujoPemex_MMUSD'] = tabla_duc['VCH_MMUSD'] - tabla_duc['Costo variable (MMUSD)'] - tabla_duc['Costos fijos (MMUSD)'] - tabla_duc['Inversiones (MMUSD)'] - tabla_duc['DERECHOS_MMUSD'] - tabla_duc['IAEEH_MMUSD'] - tabla_duc['ISR_MMUSD']
    tabla_duc['FlujoEstado_MMUSD'] = tabla_duc['DERECHOS_MMUSD'] + tabla_duc['IAEEH_MMUSD'] + tabla_duc['ISR_MMUSD']
    tabla_duc['RentaPetrolera_MMUSD'] = tabla_duc['VCH_MMUSD'] - tabla_duc['Costo variable (MMUSD)'] - tabla_duc['Costos fijos (MMUSD)'] - tabla_duc['Inversiones (MMUSD)']

    ## Calcular el valor presente neto de los flujos de cada grupo
    tabla2 = tabla_duc.groupby(['Campo', 'Categoria', 'tipo'])

    df = []


    for n, g in tabla2:
        print(n)
        x = np.npv(tasa_descuento, g.loc[:, 'FlujoPemex_MMUSD'])
        y = np.npv(tasa_descuento, g.loc[:, 'FlujoEstado_MMUSD'])
        z = np.npv(tasa_descuento, g.loc[:, 'RentaPetrolera_MMUSD'])
        a = np.sum(g.loc[:, 'FlujoPemex_MMUSD'])
        b = np.sum(g.loc[:, 'FlujoEstado_MMUSD'])
        c = np.sum(g.loc[:, 'RentaPetrolera_MMUSD'])
        d = (b / c) * 100
        e = np.sum(g.loc[:, 'Gas (mmpc)']) / 1000
        f = np.sum(g.loc[:, 'Crudo (mb)']) / 1000
        h = np.sum(g.loc[:, 'Condensado (mb)']) /1000
        i = e/5.15 + f + h
        df.append({'Campo': g.loc[0, 'Campo'], 'Categoria': g.loc[0, 'Categoria'], 'Tipo': g.loc[0, 'tipo'], 'VPN Pemex': x,
                   'Flujo Pemex': a, 'VPN Estado': y, 'Flujo Estado': b, 'VPN Renta Petrolera': z,
                   'Flujo Renta Petrolera': c, 'Government take': d, 'Vol Aceite (mmb)': f, 'Vol Gas (mmmpc)': e, 'Vol Condensado (mmb)': h, 'Vol PCE (mmb)': i})

    data = pd.DataFrame(df)
    data = data[['Campo', 'Categoria', 'Tipo', 'Vol Aceite (mmb)', 'Vol Gas (mmmpc)', 'Vol Condensado (mmb)', 'Vol PCE (mmb)', 'Flujo Renta Petrolera', 'VPN Renta Petrolera', 'Flujo Estado', 'VPN Estado', 'Flujo Pemex', 'VPN Pemex', 'Government take']]
    west = xw.Book('Plantilla_RF.xlsm').sheets['Estimaciones']
    west.range('A1').value = tabla_duc


    ## Publicar los resultados en la hoja de Excel
    wb.range('B30').value = round(float(data.loc[0, 'Vol Aceite (mmb)']), 2)
    wb.range('F30').value = round(float(data.loc[0, 'Vol Gas (mmmpc)']), 2)
    wb.range('J30').value = round(float(data.loc[0, 'Vol Condensado (mmb)']), 2)


    wb.range('B34').value = round(float(data.loc[0, 'Government take']), 2)
    wb.range('G34').value = round(float(data.loc[0, 'VPN Renta Petrolera']), 2)

    wb.range('B37').value = round(float(data.loc[0, 'VPN Pemex']), 2)
    wb.range('B40').value = round(float(data.loc[0, 'Flujo Pemex']), 2)

    wb.range('G37').value = round(float(data.loc[0, 'VPN Estado']), 2)
    wb.range('G40').value = round(float(data.loc[0, 'Flujo Estado']), 2)

    tabla_duc = tabla_duc.set_index('Año')
    tabla_duc['Opex'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)']
    fig = plt.figure()
    plt.bar(tabla_duc.index, tabla_duc['VCH_MMUSD'], width = 100, color='blue', label = 'VCH')
    plt.bar(tabla_duc.index, tabla_duc['Inversiones (MMUSD)'], width = 100, color='red', label = 'Capex')
    plt.bar(tabla_duc.index, tabla_duc['Opex'], width = 100, color='g', label = 'Opex')
    plt.ylabel('millones de dolares')
    plt.legend()

    #plt.rcParams["axes.grid.axis"] = "x"    ## hacer que solo aparezcan las lineas horizontales
    #plt.rcParams["axes.grid"] = True
    ax = plt.gca()
    ax.grid(True)
    plt.grid()

    wb.pictures.add(fig, name='MyPlot', update=True, left = wb.range('A44').left, top=wb.range('A44').top)

######################################################################################################################################################################################################################################################
######################################################################################################################################################################################################################################################


elif tipo_analisis == 'Region Fiscal' and regimen=='Asignacion':

    ## Importar la informacion de reservas
    datos = xw.Book('Plantilla_RF.xlsm').sheets['Datos']
    perfiles = datos.range('A1').expand().options(pd.DataFrame).value.reset_index()
    perfiles.fillna(0, inplace=True)
    perfiles = perfiles.drop(labels = 'Total', axis=1)
    #perfiles[['Region', 'Activo', 'Categoria']] = perfiles[['Region', 'Activo', 'Categoria']].astype('category')
    perfiles['id'] = perfiles.groupby(['Campo', 'Categoria']).ngroup()


    ## Reacomodar la tabla
    perfilesm = pd.melt(perfiles, id_vars=['Region', 'Activo', 'Campo', 'Asignacion / Contrato', 'Categoria', 'Perfil', 'id'], var_name='Año', value_name='Monto')
    perfilesm = perfilesm.groupby(['Region', 'Activo', 'Campo', 'Asignacion / Contrato', 'Categoria', 'id', 'Año', 'Perfil'])['Monto'].aggregate('sum').unstack(level = 'Perfil').reset_index()
    tipo_campo = perfilesm.groupby(['Campo', 'Categoria'])['Crudo (mb)'].aggregate('sum').reset_index()
    tipo_campo['tipo'] = tipo_campo['Crudo (mb)'].map(lambda x: 'Asociado' if x>0 else 'No Asociado')
    tipo_campo = tipo_campo[['Campo', 'Categoria', 'tipo']]
    perfilesm = perfilesm.merge(tipo_campo, how='outer', left_on = ['Campo', 'Categoria'], right_on=['Campo', 'Categoria'])
    perfilesm['Campo'] = map(lambda x: x.upper(), perfilesm['Campo'])
    perfilesm['Campo'] = perfilesm.Campo.str.normalize('NFKD').str.encode('utf-8').str.decode('ascii', 'ignore')
    perfilesm.columns = perfilesm.columns.str.replace('ó', 'o')

    ## Unir las tablas de CAT_TIPO_CAMPO con la de PERFILESM para conocer la ubicacion
    catalogo = xw.Book('Plantilla_RF.xlsm').sheets['cat_tipo_campos']
    cat_tipo_campos = catalogo.range('A1').expand().options(pd.DataFrame).value.reset_index()
    perfilesm = perfilesm.merge(cat_tipo_campos[['CAMPO','UBICACION']], how='left', left_on='Campo', right_on='CAMPO')
    perfilesm['UBICACION'] = perfilesm['UBICACION'].replace(np.nan, 'Terrestre', regex=True)  #los campos que no fueron unidos, los identificamos como Terrestres de manera manual
    perfilesm['UBICACION'] = perfilesm['UBICACION'].replace('Aguas Someras', 'Aguas someras')

    ## Quedarnos solo con la region de interes y agrupar todos los valores
    perfilesm = perfilesm[perfilesm['Categoria'] == categoria]

    if region_fiscal == 'Aguas someras':
        mask1 = perfilesm['UBICACION'] == 'Aguas someras'
        mask3 = perfilesm['tipo'] == 'Asociado'
        perfilesm = perfilesm[mask1 & mask3]
        perfilesexpl = perfilesm
        perfilesm = perfilesm.groupby(['UBICACION', 'tipo', 'Año']).aggregate('sum').reset_index()
    elif region_fiscal == 'Terrestre':
        mask1 = perfilesm['UBICACION'] == 'Terrestre'
        mask3 = perfilesm['tipo'] == 'Asociado'
        perfilesm = perfilesm[mask1 & mask3]
        perfilesexpl = perfilesm
        perfilesm = perfilesm.groupby(['UBICACION', 'tipo', 'Año']).aggregate('sum').reset_index()
    elif region_fiscal == 'ATG':
        perfilesm['UBICACION'] = perfilesm['Activo']
        mask1 = perfilesm['UBICACION'] == 'Aceite Terciario del Golfo'
        mask3 = perfilesm['tipo'] == 'Asociado'
        perfilesm = perfilesm[mask1 & mask3]
        perfilesexpl = perfilesm
        perfilesm = perfilesm.groupby(['UBICACION', 'tipo', 'Año']).aggregate('sum').reset_index()
        perfilesm = perfilesm[perfilesm['UBICACION'] == 'Aceite Terciario del Golfo']
    elif region_fiscal == 'Aguas profundas':
        mask1 = perfilesm['UBICACION'] == 'Aguas profundas'
        mask3 = perfilesm['tipo'] == 'Asociado'
        perfilesm = perfilesm[mask1 & mask3]
        perfilesexpl = perfilesm
        perfilesm = perfilesm.groupby(['UBICACION', 'tipo', 'Año']).aggregate('sum').reset_index()
    else:
        perfilesm['UBICACION'] = perfilesm['tipo']
        mask1 = perfilesm['UBICACION'] == 'No Asociado'
        perfilesm = perfilesm[mask1]
        perfilesexpl = perfilesm
        perfilesm = perfilesm.groupby(['UBICACION', 'tipo', 'Año']).aggregate('sum').reset_index()

    ## Asignar el año como indice en formato date
    perfilesm['Año'] = pd.to_datetime((perfilesm['Año']).astype(int), format='%Y') #.dt.date.astype("datetime64[ns]")
    #perfilesm.set_index('Año', inplace=True)


    ## Crear DataFrame de precios de aceite, gas y condensados por año y unir al dataframe de perfilesm
    column_precios = pd.DataFrame(columns=['Año', 'precio_aceite', 'precio_gas', 'precio_condensado'])
    column_precios[['Año']] = pd.DataFrame(pd.date_range(perfilesm['Año'].min(), perfilesm['Año'].max(), freq='AS'))
    column_precios[['precio_aceite']] = float(wb.range('B17').value)
    column_precios[['precio_gas']] = float(wb.range('B19').value)
    column_precios[['precio_condensado']] = float(wb.range('B21').value)
    perfilesm = perfilesm.merge(column_precios, how='left', left_on='Año', right_on='Año')


    ## Variables generales
    tipo_cambio = float(wb.range('I6').value)
    pce_tc = float(wb.range('B24').value)
    pce_as = float(wb.range('B23').value)
    vh_as = float(wb.range('I20').value)
    vh_tc = float(wb.range('I21').value)
    vh_atg = float(wb.range('I22').value)
    vh_gna = float(wb.range('I24').value)
    vh_ap = float(wb.range('I23').value)

    iaeeh_expl = 1583.74
    iaeeh_ext = 6334.98
    area_km2 = float(wb.range('B15').value)

    valor_remanente = float(wb.range('I8').value)
    tasa_descuento = float(wb.range('I12').value)
    tasa_duc = 0.65
    tasa_isr = float(wb.range('I10').value)

    ##############################################################################################################
    ##                                              FUNCIONES                                                   ##
    ##############################################################################################################


    ## Estimar la tasa de DEXT
    def tasa_dext(tasa_aceite=perfilesm['precio_aceite'], tasa_gas=perfilesm['precio_gas'],tasa_condensado=perfilesm['precio_condensado']):
        pct_aceite = pd.Series(tasa_aceite).map(lambda x: 0.075 if x < 45.95 else ((x * 0.125) + 1.5) / 100).rename('tasa_aceite')
        pct_gasoc = pd.Series(tasa_gas).map(lambda x: x / 100).rename('tasa_gasoc')
        pct_gnasoc = pd.Series(tasa_gas).map(lambda x: 0 if x <= 4.79 else (((x - 5) * 0.605) / x if x < 5.5 else x / 100)).rename('tasa_gnasoc')
        pct_condensado = pd.Series(tasa_condensado).map(lambda x: 0.05 if x < 57.44 else ((0.125 * x) - 0.025) / 100).rename('tasa_condensado')
        pct_dext = pd.concat([pct_aceite, pct_gasoc, pct_gnasoc, pct_condensado], axis=1)
        return (pct_dext)


    ## Crear funcion para estimar el Derecho de Extraccion
    def dext(tabla_base=perfilesm):
        tabla_dext = pd.concat([tabla_base, tasa_dext()], axis=1)
        tabla_dext['VCH_aceite'] = tabla_dext['Crudo (mb)'] * tabla_dext['precio_aceite'] / 1000
        tabla_dext['VCH_gas'] = tabla_dext['Gas (mmpc)'] * tabla_dext['precio_gas'] / 1000
        tabla_dext['VCH_condensado'] = tabla_dext['Condensado (mb)'] * tabla_dext['precio_condensado'] / 1000
        tabla_dext['VCH_MMUSD'] = (tabla_dext['VCH_aceite'] + tabla_dext['VCH_gas'] + tabla_dext['VCH_condensado'])
        tabla_dext['dext_aceite'] = tabla_dext['VCH_aceite'] * tabla_dext['tasa_aceite']
        tabla_dext['dext_gasoc'] = (tabla_dext.VCH_gas * tabla_dext.tasa_gasoc).where(tabla_dext.tipo == 'Asociado', other=0)
        tabla_dext['dext_gnasoc'] = (tabla_dext.VCH_gas * tabla_dext.tasa_gnasoc).where(tabla_dext.tipo == 'No Asociado', other=0)
        tabla_dext['dext_condensado'] = tabla_dext.VCH_condensado * tabla_dext.tasa_condensado
        tabla_dext['DEXTH_MMUSD'] = (tabla_dext.dext_aceite + tabla_dext.dext_gasoc + tabla_dext.dext_gnasoc + tabla_dext.dext_condensado)
        tabla_dext = tabla_dext[['Año', 'DEXTH_MMUSD', 'VCH_aceite', 'VCH_gas', 'VCH_condensado', 'VCH_MMUSD']]
        return (tabla_dext)


    ## Crear funcion para estimar el Derecho de Exploracion
    dexpl_inicio = 1214.21
    dexpl_despues = 2903.54
    area_km2 = float(wb.range('B15').value)


    def dexpl(tabla_base=perfilesexpl):
        tabla_dexpl = tabla_base
        tabla_dexpl['PCEcum'] = tabla_dexpl.groupby(by=['Campo', 'Categoria'])['PCE (mb)'].cumsum()
        tabla_dexpl['etapa'] = tabla_dexpl['PCEcum'].map(lambda x: float(0) if x > 0 else float(1))
        tabla_dexpl['etapacum'] = tabla_dexpl.groupby(by=['Campo', 'Categoria'])['etapa'].cumsum()
        tabla_dexpl['etapacum'] = tabla_dexpl['etapa']*tabla_dexpl['etapacum']
        tabla_dexpl['DEXPL_MMUSD'] = tabla_dexpl['etapacum'].map(lambda x: ((dexpl_inicio*area_km2*12)/tipo_cambio)/1000000 if x<6 and x>0 else (((dexpl_despues*area_km2*12)/tipo_cambio)/1000000 if x>5 and x>0 else 0))
        tabla_dexpl = tabla_dexpl.groupby(['Año'])['DEXPL_MMUSD'].aggregate('sum').reset_index()
        tabla_dexpl = tabla_dexpl[['Año', 'DEXPL_MMUSD']]
        tabla_dexpl['Año'] = pd.to_datetime((tabla_dexpl['Año']).astype(int), format='%Y') #.dt.date.astype("datetime64[ns]")
        return (tabla_dexpl)



    ## Crear funcion para estimar el Derecho de Utilidad Compartida

    ## Escribir la funcion que defina el limite de recuperacion de costos, con base en el VCH
    def costcap_vch(tabla):
        resultado = None
        if tabla['tipo'] =='No Asociado':
            resultado = tabla['VCH_MMUSD'] * vh_gna
        else:
            if tabla['UBICACION'] == 'Aguas someras':
                resultado = tabla['VCH_MMUSD'] * vh_as
            elif tabla['UBICACION'] == 'Aguas profundas':
                resultado = tabla['VCH_MMUSD'] * vh_ap
            elif tabla['UBICACION'] == 'Terrestre':
                resultado = tabla['VCH_MMUSD'] * vh_tc
            elif tabla['UBICACION'] == 'Aceite Terciario del Golfo':
                resultado = tabla['VCH_MMUSD'] * vh_atg
        return resultado

    ## Escribir la funcion que defina el limite de recuperacion de costos, con base en PCE
    def costcap_pce(tabla):
        resultado = None
        if tabla['tipo'] == 'No Asociado':
            resultado = 0
        else:
            if tabla['UBICACION'] == 'Aguas someras':
                resultado = (tabla['PCE (mb)'] * pce_as)/1000
            elif tabla['UBICACION'] == 'Terrestre':
                resultado = (tabla['PCE (mb)'] * pce_tc)/1000
            else:
                resultado = 0
        return resultado

    def duc(tabla_base = perfilesm, dext = dext(), dexpl = dexpl()):
        deduc_exploracion = float(wb.range('I15').value)
        deduc_desarrollo = float(wb.range('I16').value)
        deduc_infraestructura = float(wb.range('I17').value)
        periodos_deduc_desarrollo = int(1/deduc_desarrollo)  #defino los periodos a deducir la inversion de desarrollo
        periodos_deduc_infraestructura = int(1/deduc_infraestructura)  #defino los periodos a deducir la inversion de infraestructura
        tabla_duc = tabla_base.merge(dext, how='left', on = ['Año']).merge(dexpl, how='left', on=['Año'])
        tabla_duc = tabla_duc[['Año', 'tipo', 'UBICACION', 'Crudo (mb)', 'Gas (mmpc)', 'Condensado (mb)', 'Costo variable (MMUSD)', 'Costos fijos (MMUSD)', 'Inversiones (MMUSD)', 'PCE (mb)', 'DEXTH_MMUSD', 'DEXPL_MMUSD', 'VCH_aceite', 'VCH_gas', 'VCH_condensado', 'VCH_MMUSD']]
        tabla_duc['Costcap_VCH'] = tabla_duc.apply(costcap_vch,axis=1)
        tabla_duc['Costcap_PCE'] = tabla_duc.apply(costcap_pce,axis=1)
        tabla_duc['Costcap_max'] = tabla_duc[['Costcap_VCH', 'Costcap_PCE']].max(axis=1)
        d = {}
        for x in range(0, periodos_deduc_desarrollo):
            d["Inversiones_{0}".format(x)] = tabla_duc.groupby(['UBICACION'])['Inversiones (MMUSD)'].shift(x)*deduc_desarrollo
            tabla_duc['Inversiones_{0}'.format(x)] = d["Inversiones_{0}".format(x)]
        if 'Infraestructura (MMUSD)' in tabla_duc.columns.values:
            infra = {}
            for x in range(0, periodos_deduc_infraestructura):
                infra["Infraestructura_{0}".format(x)] = tabla_duc.groupby(['UBICACION'])['Infraestructura (MMUSD)'].shift(x) * deduc_infraestructura
                tabla_duc['Infraestructura_{0}'.format(x)] = d["Infraestructura_{0}".format(x)]
                tabla_duc['Inversiones_Deduc'] = tabla_duc.iloc[:, tabla_duc.columns.str.startswith('Inversiones_')].sum(axis=1)
                tabla_duc['Infraestructura_Deduc'] = tabla_duc.iloc[:, tabla_duc.columns.str.startswith('Infraestructura_')].sum(axis=1)
                tabla_duc['GastoPorDeduc'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)'] + tabla_duc['Inversiones_Deduc'] + tabla_duc['Infraestructura_Deduc']
                tabla_duc['GastoDeducible'] = tabla_duc[['GastoPorDeduc', 'Costcap_max']].min(axis=1)
                return(tabla_duc)
        else:
            tabla_duc['Inversiones_Deduc'] = tabla_duc.iloc[:, tabla_duc.columns.str.startswith('Inversiones_')].sum(axis=1)
            tabla_duc['GastoPorDeduc'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)'] + tabla_duc['Inversiones_Deduc']
            tabla_duc['GastoDeducible'] = tabla_duc[['GastoPorDeduc', 'Costcap_max']].min(axis=1)
            return(tabla_duc)

    tabla_duc = duc()

    ## Creo una variable que divida los grupos de acuerdo al Campo y Categoria para aplicar el siguiente for loop
    tabla_duc2 = tabla_duc
    tabla_duc2 = tabla_duc2.groupby(['UBICACION'])

    df = pd.DataFrame(columns = tabla_duc.columns)  #Creo un data.frame que tenga los mismos nombres de columnas para pegar los calculos

    ## El siguiente for-loop itera sobre cada grupo para calcular el valor acumulado a deducir cada año
    for n, g in tabla_duc2:
        if n == tabla_duc.loc[0, 'UBICACION']:
            print(n)
            g = g.reset_index()
            g.loc[1, 'dif_costcap_invpordeduc'] = g.loc[0, 'GastoPorDeduc'] - g.loc[0, 'GastoDeducible']    # calculo la inversion no deducida en el primer año y la aviento al segundo año
            g.loc[:, 'dif_costcap_invpordeduc'].fillna(value=0, axis=0, inplace=True)
            for i in range(1, len(g['Año'])):
                g.loc[i, 'GastoPorDeduc'] = g.loc[i, 'GastoPorDeduc'] + g.loc[i, 'dif_costcap_invpordeduc']        # sumo la inversion no deducida del primer año a la inversion por deducir del año 2
                g.loc[i+1, 'dif_costcap_invpordeduc'] = g.loc[i, 'GastoPorDeduc'] - g.loc[i, 'GastoDeducible']     # obtengo la diferencia del gasto por deducir y lo que se puede deducir del año dos
            df = df.append(g)
        else:
            print(2+2)

    df['UO_MMUSD'] = df['VCH_MMUSD'] - df['DEXTH_MMUSD'] - df['GastoDeducible']
    df['DUC_MMUSD'] = tasa_duc * df['UO_MMUSD']
    df['DERECHOS_MMUSD'] = df['DUC_MMUSD'] + df['DEXTH_MMUSD'] + df['DEXPL_MMUSD']

    tabla_duc = df.drop(labels=['GastoDeducible', 'dif_costcap_invpordeduc', 'index'], axis=1)
    tabla_duc.dropna(axis=0, how='all', inplace=True)

    ## Crear funcion para calcular el IAEEH

    def iaeeh(tabla_base = perfilesm):
        tabla_iaeeh = tabla_base
        tabla_iaeeh['PCEcum'] = tabla_iaeeh.groupby(by=['UBICACION'])['PCE (mb)'].cumsum()
        tabla_iaeeh['etapa'] = tabla_iaeeh['PCEcum'].map(lambda x: int(0) if x > 0 else int(1))
        tabla_iaeeh['IAEEH_MMUSD'] = tabla_iaeeh['etapa'].map(lambda x: ((iaeeh_expl*area_km2*12)/tipo_cambio)/1000000 if x==1 else ((iaeeh_ext*area_km2*12)/tipo_cambio)/1000000)
        tabla_iaeeh.loc[(tabla_iaeeh['PCE (mb)'] == 0) & (tabla_iaeeh['PCEcum'] > 0), 'IAEEH_MMUSD'] = 0   # los periodos donde ya no se esta produciendo y no son exploratorios, no se cobra el IAEEH
        tabla_iaeeh = tabla_iaeeh[['Año', 'IAEEH_MMUSD']]
        return(tabla_iaeeh)

    ## Determinar el Impuesto sobre la Renta (ISR)
    tabla_duc = tabla_duc.merge(iaeeh(), how='left', on = ['Año'])
    tabla_duc['Deducciones'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)'] + tabla_duc['Inversiones_Deduc']
    tabla_duc['Ingreso gravable'] = tabla_duc['VCH_MMUSD'] - tabla_duc['DERECHOS_MMUSD'] - tabla_duc['Deducciones'] - tabla_duc['IAEEH_MMUSD']


    def perdidas_acumuladas(tabla=tabla_duc):
        tabla2 = tabla.groupby(['UBICACION'])
        df = pd.DataFrame(columns=tabla.columns)

        ## El siguiente for-loop itera sobre cada grupo para calcular el valor acumulado a deducir cada año
        for n, g in tabla2:
            print(n)
            g = g.reset_index()
            g.loc[0, 'Perdidas Acumuladas 1'] = np.min(pd.to_numeric((g.loc[0,'Ingreso gravable'], g.loc[0,'Ingreso gravable']-valor_remanente)))    # asigno si hay alguna inversion arrastrando por deducir
            for i in range(1, len(g['Año'])):
                if g.loc[i, 'Ingreso gravable'] < 0:
                    g.loc[i, 'Perdidas Acumuladas 1'] = g.loc[i-1, 'Perdidas Acumuladas 1'] + g.loc[i, 'Ingreso gravable']
                elif g.loc[i, 'Ingreso gravable'] < abs(g.loc[i-1, 'Perdidas Acumuladas 1']):
                    g.loc[i, 'Perdidas Acumuladas 1'] = g.loc[i-1, 'Perdidas Acumuladas 1'] + g.loc[i, 'Ingreso gravable']
                else:
                    g.loc[i, 'Perdidas Acumuladas 1'] = 0
            df = df.append(g)

        df['Perdidas Acumuladas 2'] = df.loc[:, 'Ingreso gravable'].rolling(min_periods=1, window=11).sum()
        df['Perdidas Acumuladas'] = df[['Perdidas Acumuladas 1', 'Perdidas Acumuladas 2']].max(axis=1)
        df['IGAP'] = np.where(df['Ingreso gravable'] <= 0, 0, np.where(df['Perdidas Acumuladas'] <= 0, 0, df[['Ingreso gravable', 'Perdidas Acumuladas']].min(axis=1)))
        return(df)

    tabla_duc = perdidas_acumuladas()

    tabla_duc['ISR_MMUSD'] = tabla_duc['IGAP'] * tasa_isr
    tabla_duc['FlujoEfectivo_MMUSD'] = tabla_duc['Ingreso gravable'] - tabla_duc['ISR_MMUSD']
    tabla_duc['FlujoPemex_MMUSD'] = tabla_duc['VCH_MMUSD'] - tabla_duc['Costo variable (MMUSD)'] - tabla_duc['Costos fijos (MMUSD)'] - tabla_duc['Inversiones (MMUSD)'] - tabla_duc['DERECHOS_MMUSD'] - tabla_duc['IAEEH_MMUSD'] - tabla_duc['ISR_MMUSD']
    tabla_duc['FlujoEstado_MMUSD'] = tabla_duc['DERECHOS_MMUSD'] + tabla_duc['IAEEH_MMUSD'] + tabla_duc['ISR_MMUSD']
    tabla_duc['RentaPetrolera_MMUSD'] = tabla_duc['VCH_MMUSD'] - tabla_duc['Costo variable (MMUSD)'] - tabla_duc['Costos fijos (MMUSD)'] - tabla_duc['Inversiones (MMUSD)']

    ## Calcular el valor presente neto de los flujos de cada grupo
    tabla2 = tabla_duc.groupby(['UBICACION'])

    df = []


    for n, g in tabla2:
        print(n)
        x = np.npv(tasa_descuento, g.loc[:, 'FlujoPemex_MMUSD'])
        y = np.npv(tasa_descuento, g.loc[:, 'FlujoEstado_MMUSD'])
        z = np.npv(tasa_descuento, g.loc[:, 'RentaPetrolera_MMUSD'])
        a = np.sum(g.loc[:, 'FlujoPemex_MMUSD'])
        b = np.sum(g.loc[:, 'FlujoEstado_MMUSD'])
        c = np.sum(g.loc[:, 'RentaPetrolera_MMUSD'])
        d = (b / c) * 100
        e = np.sum(g.loc[:, 'Gas (mmpc)']) / 1000
        f = np.sum(g.loc[:, 'Crudo (mb)']) / 1000
        h = np.sum(g.loc[:, 'Condensado (mb)']) /1000
        i = e/5.15 + f + h
        df.append({'Ubicacion': g.loc[0, 'UBICACION'], 'Categoria': categoria, 'Tipo': g.loc[0, 'tipo'], 'VPN Pemex': x,
                   'Flujo Pemex': a, 'VPN Estado': y, 'Flujo Estado': b, 'VPN Renta Petrolera': z,
                   'Flujo Renta Petrolera': c, 'Government take': d, 'Vol Aceite (mmb)': f, 'Vol Gas (mmmpc)': e, 'Vol Condensado (mmb)': h, 'Vol PCE (mmb)': i})

    data = pd.DataFrame(df)
    data = data[['Ubicacion', 'Tipo', 'Categoria', 'Vol Aceite (mmb)', 'Vol Gas (mmmpc)', 'Vol Condensado (mmb)', 'Vol PCE (mmb)', 'Flujo Renta Petrolera', 'VPN Renta Petrolera', 'Flujo Estado', 'VPN Estado', 'Flujo Pemex', 'VPN Pemex', 'Government take']]
    west = xw.Book('Plantilla_RF.xlsm').sheets['Estimaciones']
    west.range('A1').value = tabla_duc


    ## Publicar los resultados en la hoja de Excel
    wb.range('B30').value = round(float(data.loc[0, 'Vol Aceite (mmb)']), 2)
    wb.range('F30').value = round(float(data.loc[0, 'Vol Gas (mmmpc)']), 2)
    wb.range('J30').value = round(float(data.loc[0, 'Vol Condensado (mmb)']), 2)


    wb.range('B34').value = round(float(data.loc[0, 'Government take']), 2)
    wb.range('G34').value = round(float(data.loc[0, 'VPN Renta Petrolera']), 2)

    wb.range('B37').value = round(float(data.loc[0, 'VPN Pemex']), 2)
    wb.range('B40').value = round(float(data.loc[0, 'Flujo Pemex']), 2)

    wb.range('G37').value = round(float(data.loc[0, 'VPN Estado']), 2)
    wb.range('G40').value = round(float(data.loc[0, 'Flujo Estado']), 2)

    tabla_duc = tabla_duc.set_index('Año')
    tabla_duc['Opex'] = tabla_duc['Costo variable (MMUSD)'] + tabla_duc['Costos fijos (MMUSD)']
    fig = plt.figure()
    plt.bar(tabla_duc.index, tabla_duc['VCH_MMUSD'], width = 100, color='blue', label = 'VCH')
    plt.bar(tabla_duc.index, tabla_duc['Inversiones (MMUSD)'], width = 100, color='red', label = 'Capex')
    plt.bar(tabla_duc.index, tabla_duc['Opex'], width = 100, color='g', label = 'Opex')
    plt.ylabel('millones de dolares')
    plt.legend()

    #plt.rcParams["axes.grid.axis"] = "x"    ## hacer que solo aparezcan las lineas horizontales
    #plt.rcParams["axes.grid"] = True
    ax = plt.gca()
    ax.grid(True)

    wb.pictures.add(fig, name='MyPlot', update=True, left = wb.range('A44').left, top=wb.range('A44').top)


else:
    wb.range('B44').value = 'NO SE INTRODUJERON VALORES CORRECTOS'

#################################################################################################
# FIN
#################################################################################################
