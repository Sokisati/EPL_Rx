# -- coding: utf-8 --

import socket

ip = '0.0.0.0'  
port = 12345  

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

sock.bind((ip, port))

sock.listen(1)
print("baglanti bekleniyor...")

connection, address = sock.accept()
print(f"baglanti adresi: {address}")

try:
    while True:
        data = connection.recv(1024).decode()
        print(data)

except KeyboardInterrupt:
    print("kesildi.")
    connection.close()