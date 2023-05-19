import pandas as pd
from date_class import Event
import tkinter as tk
from tkinter import filedialog as fd


def Write_Event_Data(event, txt_file):  #zapiše celoten self.event_data iz Event() objekta
    global current_row
    for i in event.event_data.values():
        txt_file.write(i+'\r\n')    # SEPARATOR NEW LINE
        current_row += 1

def Check_Similarity(keys, values, index):   # preveri, če je v vesh paralelkah razpisan isti event in in teh naredi enega
    return_str = keys[index]
    skip_class = False #če mora preskočiti že obdelani razred, ki je 2 mesti naprej index + 2

    next_i = index + 1
    while next_i < len(keys)-1 and return_str[0] == keys[next_i][0] and values[index] == values[next_i]:    #če so zaporedni 8.ABC in v vseh enaki eventi npr. NRA, vrne str -> 8.ABC
        return_str += keys[next_i][-1]
        next_i += 1
    if next_i == index + 1 and return_str[0] == keys[next_i+1][0] and values[index] == values[next_i+1]:    # če se ujema dogodek v A in C razredu, v B pa ne
        return_str += keys[next_i+1][-1]
        skip_class = True
    print(return_str, next_i, skip_class)
    return (return_str, next_i, skip_class)




filename = fd.askopenfilename()

df = pd.read_excel(filename, sheet_name= "Februar")

export_file = open(r"C:\Users\TTS\Desktop\export_events.ics", "w", encoding = 'utf-8', newline = '')
print('\n   export_file.open()\n')

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
    export_file.write(data + '\r\n') # SEPARATOR NEW LINE
    current_row += 1

keys = [i for i in df.keys()]   # glava excella - [datum, dan, parni/neparni, razredi_0,..., razredi_n, opis]

for row_i,row in df.iterrows(): # index trenutne vrstice v excelu, kaj je zapisano v vrstici [enako kot zgoraj v keys, samo da so podatki celic]
    print([i for i in row])
    col_i = 3
    povecaj_i = False
    while col_i < len(keys)-1:  #pomik desno po stolpcih (vrednostih in ključih)
        # for col_i in range(3,len(keys)-1):
        if povecaj_i:   # če mora preskočiti C razred, ker ga je že dodal kot npr. 7.AC
            skip = 1
            povecaj_i = False
        else:
            skip = 0

        if not pd.isna(row[col_i]): #če vrednost v preglednici ni prazna

            razredi, novi_index, povecaj_i = Check_Similarity(keys, row, col_i) # če ima več razredov isti tip dogodka vrne npr. 7.ABC
            if len(razredi) > 3:  # če ima več razredov isti tip dogodka stori tole
                novi_event = Event(row[0], razredi, row[col_i])
                col_i = novi_index if not povecaj_i else col_i+1
            else:
                novi_event = Event(row[0], keys[col_i], row[col_i])
                col_i += 1

            novi_event.AddUID()
            print(f'--> {row[0], novi_event.event_data["SUMMARY"]}')

            # Write_Event_Data(novi_event, worksheet)
            Write_Event_Data(novi_event, export_file)
            del novi_event
        else:
            col_i += 1
        col_i += skip



export_file.write('END:VCALENDAR')
export_file.close()

print('\n   export_file.close()\n')


