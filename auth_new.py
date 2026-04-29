import os
import hashlib
import uuid
import socket
from datetime import datetime, timedelta

LICENSE_FILE = "license.lic"
USED_CODES_FILE = "used_codes.json"

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

def get_ip_last_three():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        parts = ip.split('.')
        if len(parts) == 4:
            return parts[-1]
        return "0"
    except:
        return "0"

def generate_short_identifier():
    machine_code = get_machine_code()
    machine_short = int(machine_code[-4:], 16)
    
    mac_addr = get_mac_address()
    mac_clean = mac_addr.replace(':', '').replace('-', '')
    mac_int = int(mac_clean, 16)
    
    ip_last = int(get_ip_last_three())
    
    part1 = f"{machine_short:05d}"
    part2 = f"{mac_int:015d}"
    part3 = f"{ip_last:03d}"
    
    return part1 + part2 + part3

def decode_device_identifier(identifier):
    if not identifier.isdigit():
        return None
    
    try:
        if len(identifier) >= 23:
            part1 = identifier[:5]
            part2 = identifier[5:20]
            part3 = identifier[20:23]
            
            machine_short = int(part1)
            mac_int = int(part2)
            ip_last = int(part3)
            
            machine_code = f"000000000000{format(machine_short, '04X')}"[-16:].upper()
            
            mac_hex = format(mac_int, '012X').upper()
            mac_address = ':'.join(mac_hex[i:i+2] for i in range(0, 12, 2))
            
            ip_address = f"0.0.0.{ip_last}"
            
            return {
                'machine_code': machine_code,
                'mac_address': mac_address,
                'ip_address': ip_address,
                'machine_short': machine_short
            }
        elif len(identifier) >= 20:
            machine_short = int(identifier[:5])
            mac_int = int(identifier[5:17])
            ip_last = int(identifier[17:20])
            
            machine_code = f"000000000000{format(machine_short, '04X')}"[-16:].upper()
            
            mac_hex = format(mac_int, '012X').upper()
            mac_address = ':'.join(mac_hex[i:i+2] for i in range(0, 12, 2))
            
            ip_address = f"0.0.0.{ip_last}"
            
            return {
                'machine_code': machine_code,
                'mac_address': mac_address,
                'ip_address': ip_address,
                'machine_short': machine_short
            }
        else:
            return None
    except:
        return None

def date_to_num(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.toordinal()
    except:
        return 0

def num_to_date(num):
    try:
        return datetime.fromordinal(num).strftime("%Y-%m-%d")
    except:
        return None

def generate_activate_code(identifier, expire_date):
    try:
        date_num = date_to_num(expire_date)
        if date_num == 0:
            return None
        
        identifier_hash = hashlib.sha256(identifier.encode()).hexdigest()[:8]
        identifier_int = int(identifier_hash, 16)
        
        encoded = (identifier_int ^ date_num) % 10000000000
        
        return f"{encoded:010d}"
    except:
        return None

def verify_activate_code(identifier, activate_code):
    try:
        if len(activate_code) != 10 or not activate_code.isdigit():
            return None
        
        activate_num = int(activate_code)
        
        identifier_hash = hashlib.sha256(identifier.encode()).hexdigest()[:8]
        identifier_int = int(identifier_hash, 16)
        
        decoded_date_num = activate_num ^ identifier_int
        
        expire_date = num_to_date(decoded_date_num)
        
        if expire_date:
            return expire_date
        return None
    except:
        return None

def check_license():
    if not os.path.exists(LICENSE_FILE):
        return "inactive", 0, "未激活"
    
    try:
        with open(LICENSE_FILE, 'r') as f:
            lines = f.readlines()
        
        if len(lines) >= 2:
            expire_date_str = lines[0].strip()
            identifier = lines[1].strip()
            
            local_identifier = generate_short_identifier()
            
            if identifier != local_identifier:
                return "mismatch", 0, "设备不匹配"
            
            expire = datetime.strptime(expire_date_str, "%Y-%m-%d")
            now = datetime.now()
            remain = (expire - now).days + 1
            
            if now > expire:
                return "expired", 0, "已过期"
            elif remain <= 3:
                return "warn", remain, expire.strftime("%Y-%m-%d")
            else:
                return "normal", remain, expire.strftime("%Y-%m-%d")
        
        return "inactive", 0, "未激活"
    except Exception as e:
        return "error", 0, str(e)

def is_license_valid():
    status, days, expire = check_license()
    return status in ["normal", "warn"]

def activate_new(activate_code):
    try:
        import json
        
        used_codes = {}
        if os.path.exists(USED_CODES_FILE):
            with open(USED_CODES_FILE, 'r') as f:
                used_codes = json.load(f)
        
        if activate_code in used_codes:
            return False, "激活码已被使用"
        
        identifier = generate_short_identifier()
        
        expire_date = verify_activate_code(identifier, activate_code)
        
        if not expire_date:
            return False, "激活码无效"
        
        with open(LICENSE_FILE, 'w') as f:
            f.write(f"{expire_date}\n{identifier}")
        
        used_codes[activate_code] = {
            'identifier': identifier,
            'expire_date': expire_date,
            'used_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        with open(USED_CODES_FILE, 'w') as f:
            json.dump(used_codes, f, indent=2)
        
        return True, "激活成功"
    
    except Exception as e:
        return False, str(e)