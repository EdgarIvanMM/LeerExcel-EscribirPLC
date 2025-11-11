import socket

def test_port(host, port, timeout=5):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        result = sock.connect_ex((host, port))
        sock.close()
        return result == 0
    except:
        return False

plc_ip = "192.168.0.12"
port = 44818

print(f"🔍 Verificando puerto {port} en {plc_ip}...")
if test_port(plc_ip, port):
    print(f"✅ Puerto {port} está ABIERTO")
else:
    print(f"❌ Puerto {port} está CERRADO o bloqueado")