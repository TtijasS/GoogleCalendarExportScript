import datetime
import pandas as pd
import hashlib


class Event():
    def __init__(self, datum_dogodka, razred, opis):
        self.event_data = {
            'BEGIN': 'BEGIN:VEVENT',
            'DTSTART': 'DTSTART;VALUE=DATE:',
            'DTEND': 'DTEND;VALUE=DATE:',
            'DTSTAMP': 'DTSTAMP:',
            'UID': 'UID:',
            'CREATED': 'CREATED:',
            'DESCRIPTION': 'DESCRIPTION:',
            'LAST-MODIFIED': 'LAST-MODIFIED:',
            'LOCATION': 'LOCATION:',
            'SEQUENCE': 'SEQUENCE:0',
            'STATUS': 'STATUS:CONFIRMED',
            'SUMMARY': 'SUMMARY:',
            'TRANSP': 'TRANSP:TRANSPARENT',
            'END': 'END:VEVENT'}

        self.AddSTAMPS()  # init
        self.AddDTSTART_END(datum_dogodka)
        self.AddSUMMARY(razred, opis)

    def AddDTSTART_END(self, date):
        self.event_data['DTSTART'] += date.strftime("%Y%m%d")
        self.event_data['DTEND'] += (date + pd.DateOffset(1)).strftime("%Y%m%d")

    def AddSTAMPS(self, Timestamp_now = pd.Timestamp.now()):
        self.event_data['DTSTAMP'] += Timestamp_now.strftime("%Y%m%dT121537Z")
        self.event_data['CREATED'] += Timestamp_now.strftime("%Y%m%dT%H%M%SZ")
        self.event_data['LAST-MODIFIED'] += Timestamp_now.strftime("%Y%m%dT%H%M%SZ")

    def AddDESCRIPTION(self, opomba):
        self.event_data['DESCRIPTION'] += opomba

    def AddSUMMARY(self, razred, opis):
        self.event_data['SUMMARY'] += f"{razred} {opis}"

    def AddUID(self):
        uid_string = pd.Timestamp.now().isoformat() + self.event_data['SUMMARY']
        self.event_data['UID'] += hashlib.md5(uid_string.encode('utf-8')).hexdigest() + '@google.com'




# PLONKIČ
# -> kako pretvoriš datetime v string zapis
# pd.Timestamp(year=2021, month=12, day=15, hour=12).strftime("%Y%m%d")
# '20211215'
# -> in kako povečaš za 1 dan
# (pd.Timestamp(year=2021, month=12, day=15, hour=12) + pd.DateOffset(1)).strftime("%Y%m%d")
# '20211216'

# To je to kar rabš
# df['Datum'][1].strftime("%Y%m%d")
# '20210506'
