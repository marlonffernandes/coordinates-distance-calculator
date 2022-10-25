#libs
import pandas as pd
from geopy import distance
import ctypes
import sys

df = pd.read_excel('input.xlsx')

lstDistance = []
for row in df.itertuples():
    
    #iterate
    tplCoord1 = (row.Coord1Lat, row.Coord1Lng)
    tplCoord2 = (row.Coord2Lat, row.Coord2Lng)
    intDistance = int(distance.distance(tplCoord1, tplCoord2).m)
    
    lstDistance.append(intDistance)
    
    print("Coord1 -> Coord2 = " + str(intDistance) + " m")

df['Coord1->Coord2Distance(m)'] = lstDistance

#save
df.to_excel (r'input.xlsx', index = False, header=True)
print("Finished!")

MessageBox = ctypes.windll.user32.MessageBoxW
MessageBox(None, 'Finished', 'Finished!', 0)

