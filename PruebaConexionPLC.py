from pycomm3 import LogixDriver
import time

plc_ip = "192.168.0.12"

print(f"🔧 Probando diferentes métodos de conexión con {plc_ip}...")

# Método 1: Conexión básica
try:
    print("1. Probando conexión básica...")
    with LogixDriver(plc_ip) as plc:
        if plc.connected:
            print("✅ CONEXIÓN BÁSICA EXITOSA!")
            print(f"PLC: {plc.info}")
        else:
            print("❌ Conexión básica falló")
except Exception as e:
    print(f"❌ Error conexión básica: {e}")

# Método 2: Con timeout extendido
try:
    print("\n2. Probando con timeout extendido...")
    with LogixDriver(plc_ip, timeout=30) as plc:
        if plc.connected:
            print("✅ CONEXIÓN CON TIMEOUT EXITOSA!")
            print(f"PLC: {plc.info}")
        else:
            print("❌ Conexión con timeout falló")
except Exception as e:
    print(f"❌ Error con timeout: {e}")

# Método 3: Sin inicialización automática
try:
    print("\n3. Probando sin init_tags...")
    with LogixDriver(plc_ip, init_tags=False) as plc:
        if plc.connected:
            print("✅ CONEXIÓN SIN INIT_TAGS EXITOSA!")
            print(f"PLC: {plc.info}")
        else:
            print("❌ Conexión sin init_tags falló")
except Exception as e:
    print(f"❌ Error sin init_tags: {e}")

# Método 4: Conexión directa sin context manager
try:
    print("\n4. Probando conexión directa...")
    plc = LogixDriver(plc_ip)
    plc.open()
    if plc.connected:
        print("✅ CONEXIÓN DIRECTA EXITOSA!")
        print(f"PLC: {plc.info}")
        plc.close()
    else:
        print("❌ Conexión directa falló")
except Exception as e:
    print(f"❌ Error conexión directa: {e}")

# Método 5: Con route_path (para algunos PLCs)
try:
    print("\n5. Probando con route_path...")
    with LogixDriver(plc_ip, route_path=[plc_ip]) as plc:
        if plc.connected:
            print("✅ CONEXIÓN CON ROUTE_PATH EXITOSA!")
            print(f"PLC: {plc.info}")
        else:
            print("❌ Conexión con route_path falló")
except Exception as e:
    print(f"❌ Error con route_path: {e}")