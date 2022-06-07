import pandas as pd
import numpy as np
import requests
import json
from pyairtable import Table
from datetime import datetime

# Excel Dosyalarının Okunması
data = pd.read_excel("<EXCEL DOSYASI ADRESİ>", index_col = None)

## Tarihe göre filtre yapılması
# data = giris_cikis[giris_cikis["tarih"] == "2021-03-01"]
# data.reset_index(drop=True, inplace=True)

# Airtable Bağlantısı
api_key = "<API KEY>"
table = Table(api_key, '<BASE ID>', '<TABLO ID>')

# Giriş tarihi ve saatinin kombine edilmesi
for i in range(len(data)):
    if type(data.loc[i, 'giris_saat']) != float and type(data.loc[i, 'cikis_saat']) != float:
        # Tarih ve saat sütunlarının datetime formatına dönüştürülmesi
        tarih = pd.Timestamp(data.loc[i, 'tarih']).to_pydatetime()  
        giris_saati = data.loc[i, 'giris_saat']                     
        cikis_saati = data.loc[i, 'cikis_saat']
        # Tarih ve saat sütunlarının kombine edilmesi
        giris_tarihi = datetime.combine(tarih, giris_saati) 
        cikis_tarihi = datetime.combine(tarih, cikis_saati)
        # Kombine edilen yeni tarihlerin yeni bir sütuna atanması
        data.loc[i,"giris"] = giris_tarihi 
        data.loc[i,"cikis"] = cikis_tarihi       
        # Giriş ve çıkış tarihlerinin Airtable'ın okuyabileceği bir formata getirilmesi
        data.loc[i, "giris_airtable"] = str(data.loc[i,"giris"]).replace(" ","T")+".000Z"
        data.loc[i, "cikis_airtable"] = str(data.loc[i,"cikis"]).replace(" ","T")+".000Z"


# Bilgilerin Airtable'a aktarılması
# Airtable tarih sütunlarında "Include Time" açık olmalı ve "Use the same time zone (GMT) for all collaborators" seçilmelidir
for i in range(len(data)):   
    table.create({'Name': str(data.loc[i,"personel_id"]) ,
                    "Giriş Saati" : str(data.loc[i,"giris_airtable"]),   
                    "Çıkış Saati" : str(data.loc[i,"cikis_airtable"])})
