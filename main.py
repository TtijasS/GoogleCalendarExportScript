from enum import Enum
import pandas as pd

import xlsxwriter as xlw
from date_class import Event

def Write_Event_Data(event, txt_file):
    global current_row
    for i in event.event_data.values():
        txt_file.write(i+'\r\n')    # SEPARATOR NEW LINE
        current_row += 1

def Check_Similarity(keys, items, index):   # preveri, če je v vesh paralelkah razpisan isti event in in teh naredi enega
    returning_str = keys[index]
    n = 1
    while index <= len(keys)-1 and returning_str[0] == keys[index+n][0] and items[index] == items[index+n]:
        returning_str += keys[index+n][-1]
        n += 1
    return (returning_str, index+n)

df = pd.read_excel(r"C:\Users\TTS\Desktop\ical.xlsx", sheet_name= "Februar")
# workbook = xlw.Workbook(r"C:\Users\TTS\Desktop\februar.xlsx")
# worksheet = workbook.add_worksheet()
export_events = open(r"C:\Users\TTS\Desktop\export_events.ics", "w", encoding = 'utf-8', newline = '')
print('export_events = open(...)\n')

intro_data = ["BEGIN:VCALENDAR",
              "PRODID:-//Google Inc//Google Calendar 70.9054//EN",
              "VERSION:2.0",
              "CALSCALE:GREGORIAN",
              "METHOD:PUBLISH",
              "X-WR-CALNAME:OSMS",
              "X-WR-TIMEZONE:Europe/Belgrade"]

current_row = 0  # to beleži v kateri vrstici si bil nadadnje

for data in intro_data: # zapiše začetne podatke iCal formata
    # worksheet.write(current_row, 0, data)
    export_events.write(data + '\r\n') # SEPARATOR NEW LINE
    current_row += 1

keys = [i for i in df.keys()]   # glava excella - [datum, dan, parni/neparni, razredi_0,..., razredi_n, opis]

for row_i,row in df.iterrows(): # index trenutne vrstice v excelu, kaj je zapisano v vrstici [enako kot zgoraj v keys, samo da so podatki celic]
    print([i for i in row])
    col_i = 3
    while col_i < len(keys)-1:
    # for col_i in range(3,len(keys)-1):
        if not pd.isna(row[col_i]):
            razredi, novi_index = Check_Similarity(keys, row, col_i)
            if razredi != keys[col_i]:
                novi_event = Event(row[0], razredi, row[col_i])
                col_i = novi_index-1
            else:
                novi_event = Event(row[0], keys[col_i], row[col_i])

            novi_event.AddUID()
            print(f'--> {row[0], novi_event.event_data["SUMMARY"]}')

            # Write_Event_Data(novi_event, worksheet)
            Write_Event_Data(novi_event, export_events)
            del novi_event
        col_i+=1


# worksheet.write(current_row, 0, 'END:VCALENDAR')
# workbook.close()
export_events.write('END:VCALENDAR')
export_events.close()

print('\n   workbook.close()\n')


