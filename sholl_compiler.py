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

cell_names = [] # list
intersections_net = {} # dict; keys = cell id, values = data
lengths_net = {} # ditto


for f in input_files:
	print('test',f)
	cellid_get = re.findall(r"NL_([0-9]*-*[0-9]*)",f)
	cellid_convert = map(str, cellid_get)
	cellid_str = ''.join(cellid_convert)
	if cellid_str in cell_names:
		print("Caution: duplicate file names for",cellid_str,". End the program and rename the file!")\

	else:
		cell_names.append(cellid_str)
		df = pd.read_excel(f)

		###

		# get intersections

		df_intersections = df['Intersections'].astype('float')
		df_intersections.rename({'Intersections':cellid_str})
		list_intersections = df_intersections.to_list()
		intersections_net[''.join(cellid_str)] = list_intersections

		# get lengths

		df_lengths = df['Length(µm)'].astype('float')
		df_lengths.rename({'Length(µm)':cellid_str})
		list_lengths = df_lengths.to_list()
		lengths_net[''.join(cellid_str)] = list_lengths

intersections_df = pd.DataFrame(dict([ (a,pd.Series(b)) for a,b in intersections_net.items()]))
print("Cells compiled:",intersections_df.columns)

intersections_df.to_excel('intersections-output.xlsx')

lengths_df = pd.DataFrame(dict([ (a,pd.Series(b)) for a,b in lengths_net.items()]))
print("Cells compiled:",lengths_df.columns)

lengths_df.to_excel('lengths-output.xlsx')


'''
for each input file
find intersections column
copy data -> net

find length column
copy data -> net

potential issue: different column lengths. list of lists?

# after getting all data values:
go through dictionary. find longest value = longest column
for each value in the dict
take difference of that value size and longest value size and append zeros
	check if len(value) = max_len
		if no: append zero, loop back and check again
		if yes: continue

then will be left with dict containing keys = cell ids, values = data with appended zeroes
then dict -> table -> excel




'''