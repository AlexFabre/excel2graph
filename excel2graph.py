# coding: utf-8
#!/usr/bin/python

# ------------------ Librairies ------------------
import sys
import threading
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

# ------------------ Constantes ------------------
NB_ARG = 3
Echantillon = ["A","B","C","D","E","F","G","H","I","J","K","L"]
month_1_list = ["T0","T1", "T1+1", "T1+2", "T1+3", "T1+4", "T1+5", "T1+6", "T1+7", "T1+8", "T1+9"]
month_3_list = ["T0","T3", "T3+1", "T3+2", "T3+3", "T3+4", "T3+5", "T3+6", "T3+7", "T3+8", "T3+9"]
month_6_list = ["T0","T6", "T6+1", "T6+2", "T6+3", "T6+4", "T6+5", "T6+6", "T6+7", "T6+8", "T6+9"]
Total_month_list = ["T0","T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8"]
Week_start_line = {
        "T0": 10,
        "T1": 22,
        "T1+1": 35,
        "T1+2": 47,
        "T1+3": 59,
        "T1+4": 71,
        "T1+5": 83,
        "T1+6": 95,
        "T1+7": 107,
        "T1+8": 119,
        "T1+9": 131,
        "T2": 144,
        "T3": 156,
        "T3+1": 168,
        "T3+2": 180,
        "T3+3": 192,
        "T3+4": 204,
        "T3+5": 216,
        "T3+6": 228,
        "T3+7": 240,
        "T3+8": 252,
        "T3+9": 264,
        "T4": 276,
        "T5": 288,
        "T6": 300,
        "T6+1": 312,
        "T6+2": 324,
        "T6+3": 336,
        "T6+4": 348,
        "T6+5": 360,
        "T6+6": 372,
        "T6+7": 384,
        "T6+8": 396,
        "T6+9": 408,
        "T7": 420,
        "T8": 432,
    }


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

    for ech in Echantillon:
        for key, value in data.items():
            ech_vals.append(value[param][ech]) # Getting A value for T0, then B value for T0, etc...
        
        full_vals[ech] = ech_vals.copy()
        ech_vals.clear()
        
    for ech in range(len(Echantillon)):
        plt.plot(abscisse, full_vals[Echantillon[ech]], label=Echantillon[ech])

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

def browsefunc():
        filename = fd.askopenfilename(filetypes=(("xlsx files","*.xlsx"),("All files","*.*")))
        e1.insert(tk.END, filename)

def get_input():
    global i
    path = e1.get()
    param = e2.get()
    view_type = e3.get()
    i+=1
    compute(path, param, view_type, i)

def get_input_from_event(event):
    get_input()

def compute(path, param, view_type, plot_id):
    # Read excel file
    raw_data = pd.read_excel(path, sheet_name="Données",index_col=None, header=None)

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
    if len(sys.argv) > 1:
        if len(sys.argv) < (NB_ARG+1):
            print("\033[91mERROR: Missing arguments ! Arg 1 is excel path, arg 2 is parameter, and arg 3 is view type \033[0m")
            sys.exit(1)

        path = sys.argv[1]
        param = sys.argv[2]
        view_type = sys.argv[3]

        compute(path, param, view_type, 0)
    
    else:
        window = tk.Tk()
        window.title("Excel2graph")
        window.iconbitmap("./icon.ico")
        window.geometry("500x200")
        window.resizable(0, 0)

        l1 = tk.Label(window, text = "Fichier excel")
        l1.place(x = 10, y = 23)
        e1 = tk.Entry(window, bd = 5, width=15)
        e1.place(x = 140, y = 20)
        b1=tk.Button(window, text = "Chercher...", command=browsefunc)
        b1.place(x = 300, y = 20)

        l2 = tk.Label(window, text = "Paramètre")
        l2.place(x = 10, y = 63)
        e2 = tk.Entry(window, bd = 5, width=15)
        e2.place(x = 140, y = 60)
        l22 = tk.Label(window, text = "CO2, SO2L,...")
        l22.place(x = 300, y = 63)

        l3 = tk.Label(window, text = "Durée du graphique")
        l3.place(x = 10, y = 103)
        e3 = tk.Entry(window, bd = 5, width=15)
        e3.place(x = 140, y = 100)
        l32 = tk.Label(window, text = "T1+9, T3+9, T6+9 ou T0T8")
        l32.place(x = 300, y = 103)
        btn = tk.Button(window, text = "Go!", command=get_input, width=20)
        btn.place(x = 220, y = 160)

        l4 = tk.Label(window, text = "Version 2021.11.11")
        l4.place(x = 10, y = 170)

        window.bind('<Return>', get_input_from_event)
        
        window.mainloop()

    sys.exit