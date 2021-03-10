import pandas as pd
import os

# auto detecta archivos excel con extension xlsx o xlsm en la misma carpeta y los procesa todos
# os.chdir('carpeta')  # si se quiere cambiar de carpeta de trabajo
archivos = [str(x) for x in os.listdir()
            if str(x)[0] != '~' and (str(x)[-5:] == '.xlsm' or str(x)[-5:] == '.xlsx')]

# o si se quiere solamente procesar un archivo en especifico, redefinir variable archivos
archivos = ['Detalle Servicios Zona Paga v57 (15-01-2020) Macro.xlsm']


NOMBRE_HOJA = 'ZonasPagasVigentes'
SALTARSE_FILA = 1

# diccionario para renombrar columnas inicio termino
dict_col_ini_fin = {'Inicio 1': 'Inicio 1 Lab',
                    'Término 1': 'Término 1 Lab',
                    'Inicio 2': 'Inicio 2 Lab',
                    'Término 2': 'Término 2 Lab',
                    'Inicio 1.1': 'Inicio 1 Sab',
                    'Término 1.1': 'Término 1 Sab',
                    'Inicio 2.1': 'Inicio 2 Sab',
                    'Término 2.1': 'Término 2 Sab',
                    'Inicio 1.2': 'Inicio 1 Dom',
                    'Término 1.2': 'Término 1 Dom',
                    'Inicio 2.2': 'Inicio 2 Dom',
                    'Término 2.2': 'Término 2 Dom'}
col_inifin = list(dict_col_ini_fin.values())
tipo_dias = ['Lab', 'Sab', 'Dom']


# suma una media hora a una hora en forma de tupla [HH, MM]
def sumar_mh(h):
    if h[1] == 0:
        return [h[0], 30]
    else:
        return [h[0] + 1, 0]


# chequea si orden entre dos horas en forma de tuplas
def h1_menor_h2(h1, h2):
    if h1[0] < h2[0]:
        return True
    elif h1[0] == h2[0] and h1[1] < h2[1]:
        return True
    return False


# pasa hora en forma de tupla a string
def str_hora(h):
    return f'{h[0]:02d}:{h[1]:02d}'


# calcula todas las medias horas entre dos horas y devuelve una lista de strings con ellas
def mh_entremedio(h1: str, h2: str):
    # si todo esta bien, devuelve un 0 al principio
    resultado_ = [0]

    # revisar que strings sean ideales
    try:
        assert len(h1) == 5 and ':' in h1
    except AssertionError:
        print('Error en Hora 1')
        return [1]
    try:
        assert len(h2) == 5 and ':' in h2
    except AssertionError:
        print('Error en Hora 2')
        return [2]

    hi = [int(h) for h in h1.split(':')]

    # revisar que hayan números válidos
    try:
        assert 0 <= hi[0] < 24 and 0 <= hi[1] < 60
    except AssertionError:
        print('Error por un signo negativo (-) en Hora 1')
        return [-1]

    hf = [int(h) for h in h2.split(':')]

    # revisar que hayan números válidos
    try:
        assert 0 <= hf[0] < 24 and 0 <= hf[1] < 60
    except AssertionError:
        print('Error por un signo negativo (-) en Hora 2')
        return [-2]

    # redondear hacia abajo minutos de hi a 0 o 30
    if hi[1] < 30:
        hi[1] = 0
    else:
        hi[1] = 30

    # redondear hacia arriba minutos de hf a 0 o 30
    if hf[1] > 30:
        hf[0] = hf[0] + 1
        hf[1] = 0
    elif hf[1] > 0:
        hf[1] = 30

    # enlistar medias horas como strings
    while h1_menor_h2(hi, hf):
        resultado_.append(str_hora(hi))
        hi = sumar_mh(hi)

    return resultado_


def procesar_zp(archivo):
    df = pd.read_excel(archivo,
                       sheet_name=NOMBRE_HOJA,
                       skiprows=SALTARSE_FILA)

    # sacar columnas sin encabezado
    df = df.loc[:, ~df.columns.str.match("Unnamed")]

    # renombrar inicio termino
    df.rename(columns=dict_col_ini_fin, inplace=True)

    # reemplazar '-' por espacios
    df.replace(to_replace={'-': ''}, inplace=True)

    # Inicio de Operación es date
    df['Inicio de Operación'] = df['Inicio de Operación'].astype(str)
    df['Inicio de Operación'] = pd.to_datetime(df['Inicio de Operación']).dt.date

    print('Revisando celdas que contengan un error por tener un guión(-), se borrarán')
    for col in col_inifin:
        if (df[col].str.find('-') > 0).any():
            print(f"Columna: {col}, Filas: {list(df.loc[df[col].str.find('-') > 0].index)}")
            print(f"Valores: {list(df.loc[df[col].str.find('-') > 0][col])}")

    print('Revisando celdas que contengan un error por tener un punto(.), se asumirán como un (:)')
    for col in col_inifin:
        if (df[col].str.find('.') > 0).any():
            print(f"Columna: {col}, Filas: {list(df.loc[df[col].str.find('.') > 0].index)}")
            print(f"Valores: {list(df.loc[df[col].str.find('.') > 0][col])}")

    # errores de tipeo en los horarios
    df[col_inifin] = df[col_inifin].replace('-', '', inplace=False, regex=True)
    # inicio termino tienen horas con . reemplazar por :
    df[col_inifin] = df[col_inifin].replace('\.', ':', inplace=False, regex=True)

    print(('Revisando celdas que contengan un error por ser un -1 que se traduce a 1899. ' +
           'Se asumirán como medianoche'))
    for col in col_inifin:
        df[col] = df[col].astype(str).str[0:5]
        if (df[col] == '1899-').any():
            print(f"Columna: {col}, Fila {list(df.loc[df[col] == '1899-'].index)}")
            print(f"Valores: {list(df.loc[df[col] == '1899-'][col])}")
            df.loc[df[col] == '1899-', col] = '23:59'

    # filtrar solo ZP con estado Activo
    df = df.loc[df['Estado'] == 'Activa']

    # generar columnas con servicios
    aux_ss = [f"S{x}" for x in range(22)]
    columnas_SS = [x for x in aux_ss if x in df.columns]

    # aqui se irán agregando las filas del resultado
    resultado = []
    # nombre de columnas del dataframe que se devuelve como resultado
    col_result = ['Cod_Parada_Usuario', 'Cod_ZP', 'SS', 'Dia', 'MH']

    # iterar filas del dataframe
    for index, row in df.iterrows():
        agregar_ss = []
        # enlistar SS en la fila
        for ss in columnas_SS:
            if not pd.isna(row[ss]) and row[ss] != '':
                agregar_ss.append(ss)

        # para cada dia-inicio-fin enlistar las medias horas
        agregar_mh_dia = {}
        for dia in tipo_dias:
            for i in [1, 2]:
                col1 = f'Inicio {i} {dia}'
                col2 = f'Término {i} {dia}'

                # revisar que hayan datos en inicio-termino correspondientes
                chequeo_col1 = (not pd.isna(row[col1]) and row[col1] != '')
                chequeo_col2 = (not pd.isna(row[col2]) and row[col2] != '')

                # caso que no hay ningun problema: enlistar medias horas
                if chequeo_col1 and chequeo_col2:
                    mediashoras_ = mh_entremedio(row[col1], row[col2])
                    if len(mediashoras_) > 0:
                        if mediashoras_[0] == 0:
                            agregar_mh_dia[dia] = mediashoras_[1:]
                        else:
                            # printear lugar del error en la funcion mh_entremedio
                            print(f"Fila {index}, {col1}: {row[col1]}, {col2}: {row[col2]}")

                elif not (chequeo_col1 or chequeo_col2):
                    pass
                elif not chequeo_col1:
                    print(f"ERROR: Fila {index} tiene {col2} pero no tiene {col1}")
                elif not chequeo_col2:
                    print(f"ERROR: Fila {index} tiene {col1} pero no tiene {col2}")

        # revisar si hay más de una parada
        usar_parada1 = (not pd.isna(row["Código Parada Usuario_1"]) and
                        row["Código Parada Usuario_1"] != '')
        usar_parada2 = (not pd.isna(row["Código Parada Usuario_2"]) and
                        row["Código Parada Usuario_2"] != '')
        if agregar_mh_dia:
            for ss in agregar_ss:
                for dia in agregar_mh_dia:
                    for mh in agregar_mh_dia[dia]:
                        if usar_parada1:
                            resultado.append([row["Código Parada Usuario_1"],
                                              row["Código ZP"], row[ss], dia, mh])
                        if usar_parada2:
                            resultado.append([row["Código Parada Usuario_2"],
                                              row["Código ZP"], row[ss], dia, mh])

    df_f = pd.DataFrame(resultado, columns=col_result)
    archivo = archivo.replace('.xlsm', '')
    archivo = archivo.replace('.xlsx', '')
    print(f'Guardando {archivo}_procesado.xlsx')
    df_f.to_excel(f'{archivo}_procesado.xlsx')
    return None


def main():
    print(f'Lista de archivos encontrados en carpeta: {archivos}')

    for archivo in archivos:
        print(f'Procesando archivo {archivo}')
        procesar_zp(archivo)

    print('Todo listo.')


if __name__ == '__main__':
    main()
