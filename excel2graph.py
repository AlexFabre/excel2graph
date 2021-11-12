# coding: utf-8
#!/usr/bin/python

# ------------------ Librairies ------------------
import sys
from tkinter.constants import TRUE
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import tkinter as tk
import tkinter.filedialog as fd

import codecs
if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

global i
i = 0

# ------------------ Fonctions ------------------

# Gives the week list to plot from user input view_type
def get_weeks_to_plot(view_type):
    if view_type == "T1+9":
        return month_1_list
    elif view_type == "T3+9":
        return month_3_list
    elif view_type == "T6+9":
        return month_6_list
    elif view_type == "T0T8":
        return Total_month_list
    else:
        print("\033[91mERROR: 3rd argument must be either 'T1+9' or 'T3+9' or 'T6+9' or 'T0T8'\033[0m")
        sys.exit(2)

# Gives row number in excel file from user input parameter
def get_param_row(raw_data, param):
    param_list = raw_data.iloc[6:7].values[0]
    i = 0
    for parameter in param_list:
        if parameter == param:
            return i
        i+=1
    print("\033[91mERROR: 2nd argument does not match any parameter : "+param+"\033[0m")
    sys.exit(2)

def get_param_list(raw_data, line):
    param_list = []
    p_list = raw_data.iloc[line-1:line].values[0]
    for parameter in p_list:
        if isinstance(parameter, str):
            param_list.append(parameter)
    return param_list

def display_param(raw_data, line):
    l = 0
    c = 0
    tk.Label(window, text = "Paramètres").place(x = 700, y = 20)
    param_list = get_param_list(raw_data, line)
    for param in param_list:
        r = tk.Checkbutton(window, text=param, onvalue=1, offvalue=0)
        r.place(x=460+c, y= 63+l)
        c+=100
        if(c == 600):
            c=0
            l+=30

def get_sample_list(raw_data, col, first_sample, nb_sample):
    sample_list = []
    first_sample_found = False
    n=0
    
    s_list = raw_data.iloc[0:, col:(col+1)]
    for sample in s_list[col]:
        if first_sample_found == True:
            sample_list.append(''.join(i for i in sample if not i.isdigit()))
            n+=1
            if n == nb_sample:
                break
            continue
        if isinstance(sample, str):
            if sample == first_sample:
                sample_list.append(''.join(i for i in sample if not i.isdigit()))
                first_sample_found = True
                n+=1

    return sample_list

def display_sample_list(raw_data, col, first_sample, nb_sample):
    global sample_list
    l=0
    c=0
    tk.Label(window, text = "Echantillons").place(x = 700, y = 180)

    if isinstance(col, str):
        col = ord(col) - ord("A")

    sample_list = get_sample_list(raw_data, col, first_sample, nb_sample)
    
    for sample in sample_list:
        r = tk.Checkbutton(window, text=sample, onvalue=1, offvalue=0)
        r.select()
        r.place(x=460+c, y= 203+l)
        c+=100
        if(c == 600):
            c=0
            l+=30

def get_week_list(raw_data, col, first_sample):
    week_list = []
    week_lines = {}
    i=0

    first_sample = ''.join(i for i in first_sample if not i.isdigit())
    
    s_list = raw_data.iloc[0:, col:(col+1)]
    for sample in s_list[col]:
        if isinstance(sample, str):
            if sample[0] == first_sample:
                s = list(sample)
                s[0] = "T"
                week_list.append(''.join(s))
                week_lines[sample] = i
        i+=1

    return week_list, week_lines

def display_week_list(raw_data, col, first_sample):
    l=0
    c=0
    tk.Label(window, text = "Semaines").place(x = 700, y = 280)

    if isinstance(col, str):
        col = ord(col) - ord("A")

    week_list, week_lines = get_week_list(raw_data, col, first_sample)
    
    for sample in week_list:
        r = tk.Checkbutton(window, text=sample, onvalue=1, offvalue=0)
        r.place(x=460+c, y= 303+l)
        c+=100
        if(c == 600):
            c=0
            l+=30


# Gives week start line in excel file from user input parameter
def get_week_line(param):
    return Week_start_line[param]

# Reads user input parameter units in file
def get_row_unit(data_param):
    week = data_param.iloc[7:8].values[0]
    return week[0]

# Extract data from a specified week
def extract_week(data_param, week):
    line = get_week_line(week)

    week = data_param.iloc[line:line+12]
    week = week.rename(index={line: "A",line+1: "B",line+2: "C",line+3: "D",line+4: "E",line+5: "F",line+6: "G",line+7: "H",line+8: "I",line+9: "J",line+10: "K",line+11: "L"})
    return week.to_dict()

def plot_view(data, param, view_type, unit, plot_id):
    abscisse = []
    full_vals = {}
    ech_vals = []

    plt.figure(plot_id)

    for key, value in data.items():
        abscisse.append(key) # Getting T0, T1, T1+1 etc...

    for ech in sample_list:
        for key, value in data.items():
            ech_vals.append(value[param][ech]) # Getting A value for T0, then B value for T0, etc...
        
        full_vals[ech] = ech_vals.copy()
        ech_vals.clear()
        
    for ech in range(len(sample_list)):
        plt.plot(abscisse, full_vals[sample_list[ech]], label=sample_list[ech])

    if view_type == "T1+9":
        plt.title("Evolution du "+param+" après 1 mois")
    elif view_type == "T3+9":
        plt.title("Evolution du "+param+" après 3 mois")
    elif view_type == "T6+9":
        plt.title("Evolution du "+param+" après 6 mois")
    elif view_type == "T0T8":
        plt.title("Evolution du "+param+" sur les "+str(len(Total_month_list)-1)+" mois")
    
    plt.legend(loc='upper right')
    plt.xlabel("Temps (mois)")
    plt.ylabel(unit)
    plt.show()

def charger_fichier():
    global raw_data
    filename = fd.askopenfilename(filetypes=(("xlsx files","*.xlsx"),("All files","*.*")))
    e1.insert(tk.END, filename)
    raw_data = pd.read_excel(filename,index_col=None, header=None)
    display_param(raw_data, int(e2.get()))
    display_sample_list(raw_data, e3.get(), e4.get(), int(e5.get()))
    display_week_list(raw_data, e3.get(), e4.get())

def get_user_choice():
    global i
    path = e1.get()
    param = e2.get()
    view_type = e3.get()
    i+=1
    compute(path, param, view_type, i)

def get_user_choice_from_event(event):
    get_user_choice()

def compute(path, param, view_type, plot_id):
    # Read excel file
    # raw_data = pd.read_excel(path, sheet_name="Données",index_col=None, header=None)

    # get_param_list(raw_data, 7)

    # Getting column num for user input parameter
    row = get_param_row(raw_data, param)

    # Reducing excel sheet to only parameter's data
    data_param = raw_data.iloc[0:, row:(row+1)]

    # Renaming serie with friendly param name
    data_param = data_param.rename(columns={row: param})

    # Getting unit
    unit = get_row_unit(data_param)

    # Getting weeks list needed
    week_list = get_weeks_to_plot(view_type)

    # Getting each week in a dictionary
    data = {}
    for week in week_list:
        data[week] = extract_week(data_param, week)

    # Affichage du graphe
    plot_view(data, param, view_type, unit, plot_id)

if __name__ == '__main__':
    global window
    window = tk.Tk()
    window.title("Excel2graph")
    window.iconbitmap("./icon.ico")
    window.geometry("1100x700")
    window.resizable(0, 0)

    l1 = tk.Label(window, text = "Fichier excel")
    l1.place(x = 20, y = 23)
    e1 = tk.Entry(window, bd = 3, width=15)
    e1.place(x = 140, y = 20)
    b1=tk.Button(window, text = "Chercher...", command=charger_fichier, width=10)
    b1.place(x = 300, y = 20)

    l2 = tk.Label(window, text = "Ligne des paramètres")
    l2.place(x = 20, y = 63)
    e2 = tk.Entry(window, bd = 3, width=2)
    e2.insert(tk.END, "7")
    e2.place(x = 255, y = 60)

    l3 = tk.Label(window, text = "Colonne des échantillons")
    l3.place(x = 20, y = 103)
    e3 = tk.Entry(window, bd = 3, width=2)
    e3.insert(tk.END, "C")
    e3.place(x = 255, y = 100)

    l4 = tk.Label(window, text = "Premier échantillon")
    l4.place(x = 20, y = 143)
    e4 = tk.Entry(window, bd = 3, width=2)
    e4.insert(tk.END, "A0")
    e4.place(x = 255, y = 140)

    l5 = tk.Label(window, text = "Nombre d'échantillons")
    l5.place(x = 20, y = 183)
    e5 = tk.Entry(window, bd = 3, width=2)
    e5.insert(tk.END, "12")
    e5.place(x = 255, y = 180)

    btn = tk.Button(window, text = "Go !", command=get_user_choice, width=20)
    btn.place(x = 700, y = 660)

    l4 = tk.Label(window, text = "Version 2021.11.11")
    l4.place(x = 20, y = 670)

    window.bind('<Return>', get_user_choice_from_event)
    
    window.mainloop()

    sys.exit