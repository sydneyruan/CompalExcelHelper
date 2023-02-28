import pandas as pd
import openpyxl
import PySimpleGUI as sg
import os.path
from itertools import chain
import math
import collections
import xlsxwriter

##### FUNCTIONS #####
# Store info in the format of: [col_index, platform, CPU_ID, graphic, memory]
def extract_file(file):
	info_lst = []
	platform_lst = []
	config_dic = {}
	for i in range(3, len(file.columns)):
		platform = formatting(file.iloc[2, i]).upper()
		info_lst.append([i, platform, formatting(file.iloc[4, i]), formatting(file.iloc[18, i]), formatting(file.iloc[20, i])])
		if platform not in platform_lst:
			platform_lst.append(platform)
		platform_lst.sort()

	for j in chain(range(0, 4), range(18, 21), range(82, 86)):
		if str(file.iloc[j, 1]) != "nan":
			config_dic[j] = file.iloc[j, 1]
	for k in range(4, 13):
		if str(file.iloc[k, 2]) != "nan":
			config_dic[k] = "CPU " + str(file.iloc[k, 2])
	for m in range(13, 18):
		if str(file.iloc[m, 2]) != "nan":
			config_dic[m] = "Panel " + str(file.iloc[m, 2])
	for n in range(57, 61):
		if str(file.iloc[n, 1]) != "nan":
			config_dic[n] = "3DMark 11 " + str(file.iloc[n, 1])
	return [info_lst, platform_lst, config_dic]

# get unique list of items, ignore case
def get_unique_list(container, index_search, index_reference, item_reference):
	lst = []
	for i in range(len(container)):
		item = container[i][index_search]
		item = formatting(item)
		if container[i][index_reference] == item_reference and item not in lst and item != 'nan':
			lst.append(item)
	lst.sort()
	return lst

# strip unnecesssary spaces and (\n)'s in front of/at the end of strings
def formatting(s):
	s = str(s)
	s = s.strip(' ')
	s = s.strip('\n')
	return s

# get key of a given value in a dictionary
def get_key(dic, value):
	return list(dic.keys())[list(dic.values()).index(value)]

# Handling hidden rows and columns
def ignore_hidden_columns(file, read):
	ws = openpyxl.load_workbook(file)['Result']
	hidden_cols = []
	for colLetter,colDimension in ws.column_dimensions.items():
	    if colDimension.hidden == True:
	    	colNum = 0
	    	for c in colLetter:
	    		colNum = colNum + ord(c.lower()) - 96
	    	hidden_cols.append(colNum + 1) # account for dropped header
	unhidden_col = list(set(read.columns) - set(hidden_cols))
	return read[unhidden_col]

##### FILE UPLOAD HANDLING #####
# set theme
sg.theme('DarkAmber') 

# Upload pop-up
file_upload = [
    [sg.Text("XLSX File")],
    [sg.In(size=(30, 1), enable_events=True, key="-FILE-"), sg.FileBrowse()],
]

platform_list_column = [
    [sg.Text("Platform")],
    [sg.Listbox(values=[], enable_events=True, size=(15, 20), key="-PLATFORM LIST-")],
]

cpu_list_column = [
    [sg.Text("CPU")],
    [sg.Listbox(values=[], enable_events=True, size=(13, 20), key="-CPU LIST-")],
]

graphic_list_column = [
    [sg.Text("Graphic")],
    [sg.Listbox(values=[], enable_events=True, size=(15, 20), key="-GRAPHIC LIST-")],
]

memory_list_column = [
    [sg.Text("Memory")],
    [sg.Listbox(values=[], enable_events=True, size=(10, 20), key="-MEM LIST-", select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
]

selected_list_column = [
    [sg.Text("Selected Models", size=(27, 1)), sg.Button("Remove")],
    [sg.Listbox(values=[], enable_events=True, size=(38, 20), key="-SEL LIST-")],
]

# Upload window
layout = [
	[sg.Column(file_upload, justification='center')],
	[
		sg.Column(platform_list_column),
		sg.Column(cpu_list_column),
		sg.Column(graphic_list_column),
		sg.Column(memory_list_column),
		sg.Button('>'),
		sg.VSeperator(),
		sg.Column(selected_list_column),
	],
	[sg.Text("", size=(45, 1)), sg.Button("Export", size=(6, 1)),]
]

# Create the window
window = sg.Window("Excel Helper", layout)
# list of specs for each model
sel_lst = []

# Create an event loop
while True:
    event, values = window.read()
    # End program if user closes window or presses the Confirm button
    if event == sg.WIN_CLOSED:
        break

    # A file is uploaded
    if event == "-FILE-":
    	file = values["-FILE-"]
    	read = pd.read_excel(file, sheet_name="Result", header=None, engine='openpyxl')
    	read = read.dropna(how='all', axis=0)
    	read = read.dropna(how='all', axis=1)
    	read = ignore_hidden_columns(file, read)
    	spec_fmt = read.iloc[:, :3].values.tolist()
    	info_lst, platform_lst, cfg_dic = extract_file(read)
    	window["-PLATFORM LIST-"].update(platform_lst)

    	sorted_dic = collections.OrderedDict(sorted(cfg_dic.items()))
    	cfg_lst = list(sorted_dic.values())

    # A platform is chosen from the platform list
    elif event == "-PLATFORM LIST-" and values["-PLATFORM LIST-"]:
    	# clear previous selection
    	cpu = None
    	graphic = None

    	platform = values["-PLATFORM LIST-"][0]
    	cpu_lst = get_unique_list(info_lst, 2, 1, platform)
    	window["-CPU LIST-"].update(cpu_lst)
    	window["-GRAPHIC LIST-"].update("")
    	window["-MEM LIST-"].update("")

    # A CPU is chosen from the CPU List
    elif event == "-CPU LIST-" and values["-CPU LIST-"]:
    	cpu = values["-CPU LIST-"][0]
    	graphic_set1 = set(get_unique_list(info_lst, 3, 2, cpu))
    	graphic_set2 = set(get_unique_list(info_lst, 3, 1, platform))
    	graphic_lst = list(graphic_set1.intersection(graphic_set2))
    	graphic_lst.sort()
    	window["-GRAPHIC LIST-"].update(graphic_lst)
    	window["-MEM LIST-"].update("")

    # A graphic is chosen from the Graphic List
    elif event == "-GRAPHIC LIST-" and values["-GRAPHIC LIST-"]:
    	graphic = values["-GRAPHIC LIST-"][0]
    	mem_set1 = set(get_unique_list(info_lst, 4, 3, graphic))
    	mem_set2 = set(get_unique_list(info_lst, 4, 2, cpu))
    	mem_set3 = set(get_unique_list(info_lst, 4, 1, platform))
    	mem_lst = list(mem_set1.intersection(mem_set2.intersection(mem_set3)))
    	mem_lst.sort()
    	window["-MEM LIST-"].update(mem_lst)

    # A memory is chosen from the Memory List, user exports selection
    elif event == ">" and values["-MEM LIST-"]:
    	for mem in values["-MEM LIST-"]:
    		selection = ' > '.join([platform, cpu, graphic, mem])
    		if selection not in sel_lst:
    			sel_lst.append(selection)
    			sel_lst.sort()
    			window["-SEL LIST-"].update(sel_lst)

    # User clicks remove button
    elif event == "Remove" and values["-SEL LIST-"]:
    	remove_item = values["-SEL LIST-"][0]
    	sel_lst.remove(remove_item)
    	window["-SEL LIST-"].update(sel_lst)

    # A config is chosen from the Config List
    elif event == "Export" and sel_lst:
    	table = []
    	for key in sorted_dic.keys():
    		line = spec_fmt[key][:4]
    		print(line)
    		for sel in sel_lst:
    			pf, cpu, gf, mem = [formatting(item) for item in sel.split(">")]
    			for item in info_lst:
    				if item[1] == pf and item[2] == cpu and item[3] == gf and item[4] == mem:
    					r = read.iloc[key, item[0]]
    					if key == 0 and type(r) != str and math.isnan(r):
    						i = 1
    						while type(r) != str and math.isnan(r):
    							r = read.iloc[key, item[0]-i]
    							i = i + 1
    					line.append(r)
    					break
    		if line not in table and line:
    			table.append(line)

    	df = pd.DataFrame(table)
    	while True:
    		try:
    			df.to_excel("output.xlsx", header=None, index=None)
    		except xlsxwriter.exceptions.FileCreateError as e:
    			sg.popup("Please close \"output.xlsx\" if it is open in Excel.\n"+
                             "Press \"OK\" after the file is closed.", title = "Error")
    			continue
    		break
    	sg.popup("Succesfully exported. Check \"output.xlsx\" in your current directory.", title = "Successful")

window.close()