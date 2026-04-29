import uuid
import socket

def get_machine_code():
    try:
        return str(uuid.getnode()).zfill(16)[:16]
    except:
        return "0000000000000000"

def get_mac_address():
    try:
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) 
                       for elements in range(0,2*6,2)][::-1])
        return mac.upper()
    except:
        return "00:00:00:00:00:00"

def get_ip_address():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "0.0.0.0"

print("=" * 40)
print("本机设备信息")
print("=" * 40)
print(f"机器码: {get_machine_code()}")
print(f"机器码后四位: {get_machine_code()[-4:]}")
print(f"MAC地址: {get_mac_address()}")
print(f"IP地址: {get_ip_address()}")
print(f"IP后三位: {get_ip_address().split('.')[-1]}")
print("=" * 40)