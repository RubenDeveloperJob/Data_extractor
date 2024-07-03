import uuid
import psutil
import socket
import subprocess
from pyad import aduser
import os
from datetime import datetime
import win32api


def get_mac_address():
    mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff)
                    for elements in range(0, 2 * 6, 2)][::-1])
    return mac


def get_wifi_mac_address():
    # Ejecutar el comando netsh wlan show interfaces
    result = subprocess.run(['netsh', 'wlan', 'show', 'interfaces'], capture_output=True, text=True)
    output = result.stdout
    return output


def get_ip_address():
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)
    return ip_address


def get_connection_info():
    connection_info = os.popen('ipconfig' if os.name == 'nt' else 'ifconfig').read()
    return connection_info


def write_to_file(data):
    with open('connection_info.txt', 'a') as file:
        file.write(data + "\n")


def get_bios_serial_number():
    # Ejecutar el comando wmic bios get serialnumber
    result = subprocess.run(['wmic', 'bios', 'get', 'serialnumber'], capture_output=True, text=True)
    output = result.stdout.strip()
    return output


def get_full_username_ad():
    try:
        # Obtener el nombre de usuario de la variable de entorno
        username = os.getenv('USERNAME')

        if username:
            # Conectar con el usuario de Active Directory
            user = aduser.ADUser.from_cn(username)

            # Obtener el nombre completo del usuario
            full_name = user.get_attribute('displayName')
            return full_name
        else:
            return None
    except Exception as e:
        print(f"Error al obtener el nombre completo del usuario desde AD: {e}")
        return None


def get_current_username_env():
    return os.getenv('USERNAME')


def main():
    mac_address = get_mac_address()
    wifi_mac_address = get_wifi_mac_address()
    ip_address = get_ip_address()
    connection_info = get_connection_info()
    sn_info = get_bios_serial_number()
    full_username_ad = get_full_username_ad()
    name=get_current_username_env()

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data = f"Timestamp: {current_time}\n{sn_info}\nNom: {full_username_ad} DNI: {name}\nLAN MAC Address: {mac_address}\nIP Address: {ip_address}\nWIFI MAC Address: {wifi_mac_address}\nINFORMACIÓN DE CONEXIÓN:\n{connection_info}\n{'-' * 40}\n"

    write_to_file(data)
    print("Data has been written to connection_info.txt")


if __name__ == "__main__":
    main()
