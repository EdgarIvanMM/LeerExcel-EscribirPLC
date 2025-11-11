from pycomm3 import LogixDriver
import pandas as pd
import os

excel_name = "prueba.xlsx"
excel_path = os.path.join(os.path.dirname(__file__), excel_name)

plc_ip = "192.168.0.12"

# Leer todas las hojas del archivo
sheets = pd.read_excel(excel_path, sheet_name=None)
print(f"Se cargaron las siguientes hojas: {list(sheets.keys())}")

# Conectar con el PLC
with LogixDriver(plc_ip) as plc:
    if plc.connected:
        print("Conexión establecida con el PLC.")

        # Recorrer todas las hojas
        for hoja, df in sheets.items():
            print(f"\nProcesando hoja: {hoja}")
            print("Columnas detectadas:", list(df.columns))

            # PRIMERO: Procesar variables por modelo (solo de filas que tengan datos)
            print("\n=== PROCESANDO VARIABLES POR MODELO ===")
            modelos_procesados = {}
            
            for index, row in df.iterrows():
                try:
                    modelo = int(row["MODELO"])
                    
                    # Solo procesar si el modelo no ha sido procesado Y si tiene datos en las nuevas columnas
                    if (modelo not in modelos_procesados and 
                        pd.notna(row["CAJA EN X "]) and  # Nota el espacio extra
                        pd.notna(row["CAJA EN Y"]) and
                        pd.notna(row["TARIMA EN X"])):
                        
                        print(f"\n--- PROCESANDO VARIABLES GLOBALES DEL MODELO {modelo} ---")
                        
                        # Leer variables por modelo
                        caja_x = int(round(float(row["CAJA EN X "]) * 10))  # ¡Espacio extra en el nombre!
                        caja_y = int(round(float(row["CAJA EN Y"]) * 10))
                        caja_z = int(round(float(row["CAJA EN Z"]) * 10))
                        tarima_x = int(round(float(row["TARIMA EN X"]) * 10))
                        tarima_y = int(round(float(row["TARIMA EN Y"]) * 10))
                        tarima_z = int(round(float(row["TARIMA EN Z"]) * 10))
                        camas_totales = int(row["CAMAS TOTALES"])
                        cajas_por_cama = int(row["CAJAS POR CAMA"])

                        # Mostrar valores por modelo
                        print(f"[Excel] Modelo {modelo} - Dimensiones Caja: X:{caja_x} Y:{caja_y} Z:{caja_z}")
                        print(f"[Excel] Modelo {modelo} - Dimensiones Tarima: X:{tarima_x} Y:{tarima_y} Z:{tarima_z}")
                        print(f"[Excel] Modelo {modelo} - Configuración: Camas:{camas_totales} CajasxCama:{cajas_por_cama}")

                        # Base del tag por MODELO (no por posición)
                        base_tag_modelo = f'RECETA_PRODUCTOS[{modelo}]'

                        # ESCRITURA AL PLC - Variables por modelo
                        plc.write(f'{base_tag_modelo}.DIMENSION_PRODUCTO.DIMENSION_X', caja_x)
                        plc.write(f'{base_tag_modelo}.DIMENSION_PRODUCTO.DIMENSION_Y', caja_y)
                        plc.write(f'{base_tag_modelo}.DIMENSION_PRODUCTO.DIMENSION_Z', caja_z)
                        plc.write(f'{base_tag_modelo}.DIMENSION_TARIMA.DIMENSION_X', tarima_x)
                        plc.write(f'{base_tag_modelo}.DIMENSION_TARIMA.DIMENSION_Y', tarima_y)
                        plc.write(f'{base_tag_modelo}.DIMENSION_TARIMA.DIMENSION_Z', tarima_z)
                        plc.write(f'{base_tag_modelo}.CANTIDAD_TOTAL_CAMAS', camas_totales)
                        plc.write(f'{base_tag_modelo}.CANTIDAD_TOTAL_POSICIONES', cajas_por_cama)

                        # VERIFICACIÓN - Variables por modelo
                        val_caja_x = plc.read(f'{base_tag_modelo}.DIMENSION_PRODUCTO.DIMENSION_X').value
                        val_caja_y = plc.read(f'{base_tag_modelo}.DIMENSION_PRODUCTO.DIMENSION_Y').value
                        val_caja_z = plc.read(f'{base_tag_modelo}.DIMENSION_PRODUCTO.DIMENSION_Z').value
                        val_tarima_x = plc.read(f'{base_tag_modelo}.DIMENSION_TARIMA.DIMENSION_X').value
                        val_tarima_y = plc.read(f'{base_tag_modelo}.DIMENSION_TARIMA.DIMENSION_Y').value
                        val_tarima_z = plc.read(f'{base_tag_modelo}.DIMENSION_TARIMA.DIMENSION_Z').value
                        val_camas = plc.read(f'{base_tag_modelo}.CANTIDAD_TOTAL_CAMAS').value
                        val_posiciones = plc.read(f'{base_tag_modelo}.CANTIDAD_TOTAL_POSICIONES').value

                        print(f"[PLC] Modelo {modelo} - Dimensiones Caja: X:{val_caja_x} Y:{val_caja_y} Z:{val_caja_z}")
                        print(f"[PLC] Modelo {modelo} - Dimensiones Tarima: X:{val_tarima_x} Y:{val_tarima_y} Z:{val_tarima_z}")
                        print(f"[PLC] Modelo {modelo} - Configuración: Camas:{val_camas} CajasxCama:{val_posiciones}")

                        # Marcar modelo como procesado
                        modelos_procesados[modelo] = True
                        print(f"✓ Variables globales del Modelo {modelo} escritas correctamente")
                        
                except Exception as e:
                    # Si hay error, continuar con la siguiente fila (puede ser fila vacía)
                    continue

            # SEGUNDO: Procesar variables por posición (todas las filas)
            print(f"\n=== PROCESANDO VARIABLES POR POSICIÓN ===")
            for index, row in df.iterrows():
                try:
                    modelo = int(row["MODELO"])
                    dejada = int(row["DEJADA"])
                    cama = int(row["CAMA"])

                    # VARIABLES POR POSICIÓN (existentes)
                    # Usar fillna(0) para manejar celdas vacías
                    pos_x = int(round(float(row["POSICION X"]) * 10))
                    pos_y = int(round(float(row["POSICION Y"]) * 10))
                    giro = int(row["GIRO"])
                    cajas = int(row["CAJAS"])
                    agarre = int(row["AGARRE"])

                    # PROCESAR VARIABLES POR POSICIÓN
                    print(f"[Excel] Modelo {modelo}, Cama {cama}, Dejada {dejada} -> X:{pos_x} Y:{pos_y} Giro:{giro} Cajas:{cajas} Agarre:{agarre}")

                    # Base del tag por POSICIÓN
                    base_tag_posicion = f'RECETA_PRODUCTOS[{modelo}].PALE.CAMA[{cama}].POSICION[{dejada}]'

                    # Escritura al PLC - Variables por posición
                    plc.write(f'{base_tag_posicion}.PROC_POS_X_INT', pos_x)
                    plc.write(f'{base_tag_posicion}.PROC_POS_Y_INT', pos_y)
                    plc.write(f'{base_tag_posicion}.PROC_GIRO', giro)
                    plc.write(f'{base_tag_posicion}.PROC_CANTIDAD_AGARRE', cajas)
                    plc.write(f'{base_tag_posicion}.PROC_TIPO_AGARRE', agarre)

                    # Leer de vuelta para verificar
                    val_x = plc.read(f'{base_tag_posicion}.PROC_POS_X_INT').value
                    val_y = plc.read(f'{base_tag_posicion}.PROC_POS_Y_INT').value
                    val_giro = plc.read(f'{base_tag_posicion}.PROC_GIRO').value
                    val_cajas = plc.read(f'{base_tag_posicion}.PROC_CANTIDAD_AGARRE').value
                    val_agarre = plc.read(f'{base_tag_posicion}.PROC_TIPO_AGARRE').value

                    # Mostrar los valores escritos/verificados
                    print(f"[PLC]   Modelo {modelo}, Cama {cama}, Dejada {dejada} -> X:{val_x} Y:{val_y} Giro:{val_giro} Cajas:{val_cajas} Agarre:{val_agarre}")
                    
                    # Verificar si coinciden
                    if val_x == pos_x and val_y == pos_y:
                        print(f"✓ Posición escrita correctamente")
                    else:
                        print(f"✗ Discrepancia en valores de posición")

                except Exception as e:
                    print(f"Error en hoja {hoja}, fila {index+2}: {e}")

    else:
        print("No se pudo conectar con el PLC.")