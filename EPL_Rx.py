# -- coding: utf-8 --

import socket
import json
import time
import xlwings as xw
import os
from datetime import datetime

wb = xw.Book(r'C:\EPL_Rx\Yer_Istasyonu_veri.xlsx')
sheet = wb.sheets['Sheet1']

header_data = [
    'takimNo', 'Veri paket No', 'Time', 'Pressure', 'High', 'Speed', 'Temperature',
    'pilGerilimi', 'gpsLat', 'gpsLong', 'gpsAlt', 'pitch', 'roll', 'yaw', 'donusHizi'
]

sheet.range('A1').value = header_data

class GonderilecekVeriler:
    def __init__(self, takimNo, veriPaketNo, gondermeSaatiVeTarih, basinc, yukseklik, inisHizi, sicaklik, pilGerilimi, gpsLat, gpsLong, gpsAlt, pitch, roll, yaw, donusHizi):
        self.takimNo = takimNo
        self.veriPaketNo = veriPaketNo
        self.gondermeSaatiVeTarih = gondermeSaatiVeTarih
        self.yukseklik = yukseklik
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

        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
        
        formatted_date = datetime.strptime(alinan_veri.gondermeSaatiVeTarih, '%Y-%m-%d %H:%M:%S')
        
        formatted_date_str = formatted_date.strftime('%Y-%m-%d %H:%M:%S')

        
        row_data = [
            alinan_veri.takimNo, alinan_veri.veriPaketNo, formatted_date, alinan_veri.basinc,
            alinan_veri.yukseklik, alinan_veri.inisHizi, alinan_veri.sicaklik, alinan_veri.pilGerilimi,
            alinan_veri.gpsLat, alinan_veri.gpsLong, alinan_veri.gpsAlt, alinan_veri.pitch, alinan_veri.roll,
            alinan_veri.yaw, alinan_veri.donusHizi
        ]

        sheet.range('A' + str(last_row)).value = row_data

        print(f"Alinan veri no: {alinan_veri.veriPaketNo}")
        print(f"Sicaklik: {alinan_veri.sicaklik}")
        print(f"Pil Gerilimi: {alinan_veri.pilGerilimi}")
        print()

        time.sleep(1)

# CTRL+C
except KeyboardInterrupt:
    print("Kesildi.")
    connection.close()