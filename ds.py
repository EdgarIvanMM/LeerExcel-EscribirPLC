from pycomm3 import LogixDriver
import pandas as pd

PLC_IP = '192.168.0.12'
excel_file = 'prueba.xlsx'  # nombre de archivo

# Leer todas las hojas del archivo Excel
try:
    hojas = pd.read_excel(excel_file, sheet_name=None, header=0)
    print(f"\n Se cargaron las siguientes hojas: {list(hojas.keys())}\n")
except Exception as e:
    print(f'Error leyendo el archivo Excel: {e}')
    exit()

# Conectar al PLC y escribir datos
try:
    with LogixDriver(PLC_IP) as plc:
        if not plc.connected:
            print(f'No se pudo conectar al PLC en {PLC_IP}')
            exit()

        print(f'\nConexión establecida con el PLC en {PLC_IP}\n')

        for nombre_hoja, df in hojas.items():
            print(f'\nProcesando hoja: {nombre_hoja}')

            for i, row in df.iterrows():
                try:
                    modelo = int(row['MODELO'])
                    dejada = int(row['DEJADA'])
                    cama = int(row['CAMA'])
                    pos_x = float(row['Posición X'])
                    pos_y = float(row['Posición Y'])
                    giro = float(row['GIRO'])
                    
                    # Nuevos campos
                    cajas = int(row['CAJAS'])
                    agarre = int(row['AGARRE'])

                    # Formar tags base
                    base_tag = f'RECETA_PRODUCTOS[{modelo}].PALE.CAMA[{cama}].POSICION[{dejada}]'
                    
                    # Tags originales
                    tag_x = f'{base_tag}.PROC_POS_X'
                    tag_y = f'{base_tag}.PROC_POS_Y'
                    tag_giro = f'{base_tag}.PROC_GIRO'
                    
                    tag_cajas = f'{base_tag}.PROC_CANTIDAD_AGARRE'  
                    tag_agarre = f'{base_tag}.PROC_TIPO_AGARRE'  

                    # Escribir todos los valores
                    plc.write(
                        (tag_x, int(pos_x)), 
                        (tag_y, int(pos_y)), 
                        (tag_giro, int(giro)),
                        (tag_cajas, int(cajas)),
                        (tag_agarre, int(agarre))
                    )
                    
                    print(f'[{nombre_hoja}] Modelo {modelo} | Cama {cama} | Dejada {dejada} → X={pos_x}, Y={pos_y}, Giro={giro}, Cajas={cajas}, Agarre={agarre}')

                except Exception as write_err:
                    print(f'Error en hoja "{nombre_hoja}", fila {i+2} (modelo {row.get("MODELO")}): {write_err}')

except Exception as e:
    print(f'\nError general al conectarse o escribir al PLC: {e}')