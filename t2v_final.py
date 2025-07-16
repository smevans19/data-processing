import tkinter as tk
from tkinter import filedialog
import pandas as pd
import re
import os
import glob

## import excel file folder
root = tk.Tk()
root.withdraw()
path = filedialog.askdirectory()
input_files = glob.glob(os.path.join(path, "*.xlsx"))

## initialize outputs
cell_names = []
velocities_only = {}


for f in input_files:
    if '~' in str(f):
        print('break check')
        break

    print('test',f)
    df = pd.read_excel(f)
    cellid = re.findall(r"_0(\w+)", f) #pulls cell number from file name as list
        
    cell_names.append(cellid)
    

    # pull out all rows with 'track' and make new df
    df_tracks = df[df['Type'].str.contains('Track')] # grab all track rows
    df_tracks['Length (sum), Track Length (µm)'].astype('float')
    df_tracks.sort_values(by=['Length (sum), Track Length (µm)'], ascending = True)

    # pull out all nonzero track lengths
    df_nonzero_tracks = df_tracks[df_tracks['Length (sum), Track Length (µm)'] != 0]
    
    # delete all columns except length and n segments
    df_drop = df_nonzero_tracks[['Length (sum), Track Length (µm)','# Segments, Segment Count']]
    df_final = df_drop.rename(columns={'Length (sum), Track Length (µm)':'tlength','# Segments, Segment Count':'nsegments'})
   
    # calculate velocities
    df_final['velocities'] = (df_final['tlength'])/((1.27*(df_final['nsegments'] - 1)))

    vlist = df_final.velocities.to_list()
    velocities_only[''.join(cellid)] = vlist

velocities_df = pd.DataFrame(dict([ (a,pd.Series(b)) for a,b in velocities_only.items() ]))
print('Velocities calculated for:',velocities_df.columns)

velocities_df.to_excel('velocities-output.xlsx')
