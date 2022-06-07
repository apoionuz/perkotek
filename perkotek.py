import pandas as pd
import numpy as np
import requests
import json
from pyairtable import Table
from datetime import datetime

# Excel Dosyalarının Okunması
giris_cikis = pd.read_excel("perkotek/175000den_sonrasi.xls", index_col=None)
# personel = pd.read_excel("perkotek/personel kartları.xls", usecols=["id","kartno","adi","soyadi"])

# Tarihe göre filtre yapılması -> data1'e
data1 = giris_cikis[giris_cikis["tarih"] == "2021-03-01"]

# data1'de Index'in sıfırlanması
data1.reset_index(drop=True, inplace=True)

# Airtable Bağlantısı
api_key = "keyJT8JxQ806NtyGA"
table = Table(api_key, 'appvMr0y0mjVvQ1Bz', 'tblL8ij606tfg0gE9')

# Giriş tarihi ve saatinin kombine edilmesi
for i in range(len(data1)):
    if type(data1.loc[i, 'giris_saat']) != float:
        tarih = pd.Timestamp(data1.loc[i, 'tarih']).to_pydatetime() # "1999-08-07" verisinin datetime formatına dönüştürülmesi
        giris_saati = data1.loc[i, 'giris_saat'] 
        giris_tarihi = datetime.combine(tarih, giris_saati) # iki datetime verisinin kombine edilmesi
        data1.loc[i,"giris"] = giris_tarihi # kombine edilen verinin yeni bir sütuna atanması

# Çıkış tarihi ve saatinin kombine edilmesi
for i in range(len(data1)):
    if type(data1.loc[i, 'cikis_saat']) != float:
        tarih = pd.Timestamp(data1.loc[i, 'tarih']).to_pydatetime() 
        cikis_saati = data1.loc[i, 'cikis_saat']
        cikis_tarihi = datetime.combine(tarih, cikis_saati)
        data1.loc[i,"cikis"] = cikis_tarihi

# Giriş ve çıkış tarihlerinin Airtable'ın okuyabileceği bir şekle getirilmesi, iki yeni sütun ile
for i in range(len(data1)):
    data1.loc[i, "giris_airtable"] = str(data1.loc[i,"giris"]).replace(" ","T")+".000Z"
    data1.loc[i, "cikis_airtable"] = str(data1.loc[i,"cikis"]).replace(" ","T")+".000Z"

# Bilgilerin Airtable'a aktarılması
for i in range(len(data1)):   
    table.create({'Name': str(data1.loc[i,"personel_id"]) ,
                    "Giriş Saati" : str(data1.loc[i,"giris_airtable"]),
                    "Çıkış Saati" : str(data1.loc[i,"cikis_airtable"])})