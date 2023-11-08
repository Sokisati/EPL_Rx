# -- coding: utf-8 --

import socket
import json
import time
import xlwings as xw
import os
import numpy as np
from datetime import datetime

wb = xw.Book(r'C:\EPL_Rx\Yer_Istasyonu_veri.xlsx')
sheet = wb.sheets['Sheet1']

header_data = [
    'takimNo', 'Veri paket No', 'Time', 'Pressure', 'High', 'Speed', 'Temperature',
    'pilGerilimi', 'gpsLat', 'gpsLong', 'gpsAlt', 'pitch', 'roll', 'yaw', 'donusHizi'
]

sheet.range('A1').value = header_data

def decryption_function(key_vector, value_to_dec, data_packet_id):
    key_index = data_packet_id % len(key_vector)
    dec_value = (value_to_dec + key_vector[key_index]) / key_vector[key_index]
    return dec_value

key_vector = [16,81,33,32,5,15,13,71,43,8,31,72,4,38,71,19]

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
        try:
            data_dict = json.loads(data_json)
            alinan_veri = GonderilecekVeriler(**data_dict)
        except json.decoder.JSONDecodeError as e:
            print(f"JSON hatasi: {e}")
            continue 

        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
        

        alinan_veri.basinc = decryption_function(key_vector,alinan_veri.basinc,alinan_veri.veriPaketNo)
        alinan_veri.yukseklik = decryption_function(key_vector,alinan_veri.yukseklik,alinan_veri.veriPaketNo)
        alinan_veri.inisHizi = decryption_function(key_vector,alinan_veri.inisHizi,alinan_veri.veriPaketNo)
        alinan_veri.sicaklik = decryption_function(key_vector,alinan_veri.sicaklik,alinan_veri.veriPaketNo)
        alinan_veri.pilGerilimi = decryption_function(key_vector,alinan_veri.pilGerilimi,alinan_veri.veriPaketNo)
        alinan_veri.gpsLat = decryption_function(key_vector,alinan_veri.gpsLat,alinan_veri.veriPaketNo)
        alinan_veri.gpsLong = decryption_function(key_vector,alinan_veri.gpsLong,alinan_veri.veriPaketNo)
        alinan_veri.gpsAlt = decryption_function(key_vector,alinan_veri.gpsAlt,alinan_veri.veriPaketNo)
       

        row_data = [
            alinan_veri.takimNo, alinan_veri.veriPaketNo, alinan_veri.gondermeSaatiVeTarih, alinan_veri.basinc,
            alinan_veri.yukseklik, alinan_veri.inisHizi, alinan_veri.sicaklik, alinan_veri.pilGerilimi,
            alinan_veri.gpsLat, alinan_veri.gpsLong, alinan_veri.gpsAlt, alinan_veri.pitch, alinan_veri.roll,
            alinan_veri.yaw, alinan_veri.donusHizi
        ]

        sheet.range('A' + str(last_row)).value = row_data

        print(f"Alinan veri no: {alinan_veri.veriPaketNo}")
        print(f"Sicaklik: {alinan_veri.sicaklik}")
        print(f"Pil Gerilimi: {alinan_veri.pilGerilimi}")
        print()

        time.sleep(0.05)

# CTRL+C
except KeyboardInterrupt:
    print("Kesildi.")
    connection.close()