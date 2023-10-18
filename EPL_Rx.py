# -- coding: utf-8 --

import socket
import json
import time

class GonderilecekVeriler:
    def __init__(self, takimNo, veriPaketNo, gondermeSaatiVeTarih, basinc, inisHizi, sicaklik, pilGerilimi, gpsLat, gpsLong, gpsAlt, pitch, roll, yaw, donusHizi):
        self.takimNo = takimNo
        self.veriPaketNo = veriPaketNo
        self.gondermeSaatiVeTarih = gondermeSaatiVeTarih
        self.basinc = basinc
        self.inisHizi = inisHizi
        self.sicaklik = sicaklik
        self.pilGerilimi = pilGerilimi
        self.gpsLat = gpsLat
        self.gpsLong = gpsLong
        self.gpsAlt = gpsAlt
        self.pitch = pitch
        self.roll = roll
        self.yaw = yaw
        self.donusHizi = donusHizi

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
        data_json = connection.recv(1024).decode()
        
        data_dict = json.loads(data_json)
        alinan_veri = GonderilecekVeriler(**data_dict)

        print(f"Alinan veri no:: {alinan_veri.veriPaketNo}")
        print(f"Sicaklik: {alinan_veri.sicaklik}")
        print(f"Basinc: {alinan_veri.basinc}")
        print()
        time.sleep(1)

except KeyboardInterrupt:
    print("Kesildi.")
    connection.close()