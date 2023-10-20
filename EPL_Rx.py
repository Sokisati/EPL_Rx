# -- coding: utf-8 --

import socket
import json
import time
import xlsxwriter
import os

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
excel_file_path = os.path.join(desktop_path, "alinanVeriler.xlsx")

workbook = xlsxwriter.Workbook(excel_file_path)
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Takim No')
worksheet.write('B1', 'Veri paket no')
worksheet.write('C1', 'Gonderme tarih/saat')
worksheet.write('D1', 'Basinc')
worksheet.write('E1', 'Yukseklik')
worksheet.write('F1', 'Inis hizi')
worksheet.write('G1', 'Sicaklik')
worksheet.write('H1', 'Pil gerilimi')
worksheet.write('I1', 'GPS Lat.')
worksheet.write('J1', 'GPS Long.')
worksheet.write('K1', 'GPS Alt.')
worksheet.write('L1', 'Pitch')
worksheet.write('M1', 'Roll')
worksheet.write('N1', 'Yaw')
worksheet.write('O1', 'Donus hizi')

class GonderilecekVeriler:
    def __init__(self, takimNo, veriPaketNo, gondermeSaatiVeTarih,basinc,yukseklik, inisHizi, sicaklik, pilGerilimi, gpsLat, gpsLong, gpsAlt, pitch, roll, yaw, donusHizi):
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

        print(f"Alinan veri no:: {alinan_veri.veriPaketNo}")
        print(f"Sicaklik: {alinan_veri.sicaklik}")
        print(f"pil gerilimi: {alinan_veri.pilGerilimi}")
        print()
        columnString = str(alinan_veri.veriPaketNo + 2)
        worksheet.write('A'+columnString, str(alinan_veri.takimNo))
        worksheet.write('B'+columnString, alinan_veri.veriPaketNo)
        worksheet.write('C'+columnString, alinan_veri.gondermeSaatiVeTarih)
        worksheet.write('D'+columnString, alinan_veri.basinc)
        worksheet.write('E'+columnString, alinan_veri.yukseklik)
        worksheet.write('F'+columnString, alinan_veri.inisHizi)
        worksheet.write('G'+columnString, alinan_veri.sicaklik)
        worksheet.write('H'+columnString, alinan_veri.pilGerilimi)
        worksheet.write('I'+columnString, alinan_veri.gpsLat)
        worksheet.write('J'+columnString, alinan_veri.gpsLong)
        worksheet.write('K'+columnString, alinan_veri.gpsAlt)
        worksheet.write('L'+columnString, alinan_veri.pitch)
        worksheet.write('M'+columnString, alinan_veri.roll)
        worksheet.write('N'+columnString, alinan_veri.yaw)
        worksheet.write('O'+columnString, alinan_veri.donusHizi)
        time.sleep(1)

#CTRL+C
except KeyboardInterrupt:
    print("Kesildi.")
    connection.close()
    workbook.close()