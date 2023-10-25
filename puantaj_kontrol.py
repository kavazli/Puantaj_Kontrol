import pandas as pd
import datetime
import os
import smtplib
from email.message import EmailMessage


def csv_name():
    folder = os.listdir("C:/Users/gokhan.kaya/OneDrive - Aster Textile/Desktop/BELGELERİM/PycharmProjects/PandasProjects/puantaj_kontrol")
    csv_str = []
    for search in folder:
        if "csv" in search:
            csv_str.append(search)
    return csv_str[0]

csv_data = pd.read_csv(csv_name(), sep=";")
df = pd.DataFrame(data=csv_data)
df = df.dropna(subset="Bölüm")


data_01 = list(df["mesaitarih"])
data_02 = list(df["Giriş"])
data_03 = list(df["Çıkış"])
data_04 = list(df["OFM"])
convert_01 = []
convert_02 = []
convert_03 = []
convert_04 = []

for convert1, convert2, convert3, convert4 in zip(data_01, data_02, data_03, data_04):
    convert_01.append(datetime.datetime(int(convert1[6:10]), int(convert1[3:5]), int(convert1[0:2])))
    if type(convert2) == float:
        convert_02.append(None)
    else:
        convert_02.append(datetime.time(int(convert2[0:2]), int(convert2[3:5]), int(convert2[6:8])))
    if type(convert3) == float:
        convert_03.append(None)
    else:
        convert_03.append(datetime.time(int(convert3[0:2]), int(convert3[3:5]), int(convert3[6:8])))
    if type(convert4) == float:
        convert_04.append(None)
    else:
        convert_04.append(datetime.time(int(convert4[0:2]), int(convert4[3:5])))


df["mesaitarih"] = convert_01
df["Giriş"] = convert_02
df["Çıkış"] = convert_03
df["OFM"] = convert_04



date = list(df["mesaitarih"])
days = {0: "Haftaiçi", 1: "Haftaiçi", 2: "Haftaiçi", 3: "Haftaiçi", 4: "Haftaiçi", 5: "Cumartesi", 6: "Pazar"}
days_transfer = []
for aktar in date:
    days_transfer.append(days.get(aktar.weekday()))

df["Günler"] = days_transfer

df["Notlar"] = None

df = df.loc[:, ["sicilno", "AltFirma", "Bölüm", "mesaitarih", "Giriş", "Çıkış", "OFM", "Günler", "İzin Açıklama", "Notlar"]]




summary_filter = {1: (df["Giriş"].notna()) & (df["Çıkış"].isna()),
                  2: (df["Giriş"].isna()) & (df["Çıkış"].notna()),
                  3: (df["Günler"] == "Cumartesi") & (df["Giriş"].notna()) & (df["Çıkış"].notna()) & (df["OFM"] == datetime.time(00, 00, 00)),
                  4: (df["Günler"] == "Pazar") & (df["Giriş"].notna()) & (df["Çıkış"].notna()) & (df["OFM"] == datetime.time(00, 00, 00)),
                  5: (df["Günler"] == "Haftaiçi") & (df["Giriş"].isna()) & (df["Çıkış"].isna()) & (df["İzin Açıklama"].isna()),
                  6: (df["Günler"] == "Haftaiçi") & (df["Giriş"].notna()) & (df["Çıkış"] > datetime.time(19, 00, 00)) & (df["OFM"] == datetime.time(00, 00, 00))}

frame = []
for i in range(1, len(summary_filter)+1):
    frame.append("frame" + str(i))


tables = {}
for x in range(0, len(summary_filter)):
    tables[frame[x]] = df


tables["frame1"] = df.loc[summary_filter[1], :]
tables["frame1"].loc[:, "Notlar"] = "Çıkış bilgisi eksik"
tablo_1 = pd.DataFrame(data=tables["frame1"])


tables["frame2"] = df.loc[summary_filter[2], :]
tables["frame2"].loc[:, "Notlar"] = "Giriş Bilgisi EKsik"
tablo_2 = pd.DataFrame(data=tables["frame2"])


tables["frame3"] = df.loc[summary_filter[3], :]
tables["frame3"].loc[:, "Notlar"] = "OFM bilgisi eksik"
tablo_3 = pd.DataFrame(data=tables["frame3"])


tables["frame4"] = df.loc[summary_filter[4], :]
tables["frame4"].loc[:, "Notlar"] = "OFM bilgisi eksik"
tablo_4 = pd.DataFrame(data=tables["frame4"])


tables["frame5"] = df.loc[summary_filter[5], :]
tables["frame5"].loc[:, "Notlar"] = "İzin açıklama bilgisi eksik"
tablo_5 = pd.DataFrame(data=tables["frame5"])


tables["frame6"] = df.loc[summary_filter[6], :]
tables["frame6"].loc[:, "Notlar"] = "OFM bilgisi eksik"
tablo_6 = pd.DataFrame(data=tables["frame6"])

genel_tablo = pd.concat([tablo_1, tablo_2, tablo_3, tablo_4, tablo_5, tablo_6], axis=0)

genel_tablo.to_excel("puantaj_kontrol.xlsx", sheet_name="Günlük Paporlar")


mail_subject = "PuantajKontrol"
mail_from = "pythonmailgonderim@gmail.com"
mail_to = ["ayten.guler@astertextile.com", "serife.yaprak@astertextile.com", "ferhat.sen@astertextile.com",
           "hande.yazici@astertextile.com", "ahmet.deniz@astertextile.com"]
mail_mesaj = "Merhaba, Lütfen ekteki puantaj özet dosyasını kendi tesisiniz bazında kontrol ediniz. İyi Çalışmalar."
appPassword = "clbj ioyn gmhy pqzp"

with open("puantaj_kontrol.xlsx", "rb") as f:
    file = f.read()
    file_name = f.name

mail = EmailMessage()
mail["Subject"] = mail_subject
mail["From"] = mail_from
mail["To"] = mail_to
mail.set_content(mail_mesaj)
mail.add_attachment(file, maintype="application", subtype="octet-stream", filename=file_name)

with smtplib.SMTP_SSL("smtp.gmail.com") as sent:
    sent.login("pythonmailgonderim@gmail.com", appPassword)
    sent.send_message(mail)
    sent.quit()











