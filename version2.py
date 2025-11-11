from pycomm3 import LogixDriver
import pandas as pd

PLC_IP = '192.168.0.12'
excel_file = 'prueba.xlsx'  # Nombre de archivo de excel

# Leer la hoja con encabezados
try:
    df = pd.read_excel(excel_file, sheet_name='Hoja1', header=0)

    print("\n Contenido del archivo Excel:")
    print(df.head())

    print("\n Columnas detectadas:")
    print(df.columns)
except Exception as e:
    print(f' Error leyendo el archivo Excel: {e}')
    exit()

# Conectar al PLC y escribir datos
try:
    with LogixDriver(PLC_IP) as plc:
        if not plc.connected:
            print(f' No se pudo conectar al PLC en {PLC_IP}')
            exit()

        print(f'\n Conexión establecida con el PLC en {PLC_IP}\n')

        for i, row in df.iterrows():
            try:
                modelo = int(row['MODELO'])
                dejada = int(row['DEJADA'])
                cama = int(row['CAMA'])
                pos_x = float(row['Posición X'])
                pos_y = float(row['Posición Y'])
                giro = float(row['GIRO'])

                # Formar tags
                base_tag = f'HMI_RECETA_PRODUCTOS[{modelo}].PALE.CAMA[{cama}].DEJADA[{dejada}]'
                tag_x = f'{base_tag}.PROC_POS_X'
                tag_y = f'{base_tag}.PROC_POS_Y'
                tag_giro = f'{base_tag}.PROC_GIRO'

                # Escribir valores
                plc.write((tag_x, int(pos_x)), (tag_y, int(pos_y)), (tag_giro, int(giro)))
                print(f' Modelo {modelo} | Cama {cama} | Dejada {dejada} → X={pos_x}, Y={pos_y}, Giro={giro}')

            except Exception as write_err:
                print(f' Error en fila {i+2} (modelo {modelo}, dejada {dejada}): {write_err}')

except Exception as e:
    print(f'\n Error general al conectarse o escribir al PLC: {e}')
