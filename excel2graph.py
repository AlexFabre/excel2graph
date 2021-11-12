# coding: utf-8
#!/usr/bin/python

# ------------------ Librairies ------------------
import sys
from matplotlib import colors
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

color_list = ["firebrick","salmon","orange","burlywood","olive","yellowgreen","lightgreen","turquoise","lightskyblue","royalblue","mediumorchid","hotpink"]

# ------------------ Fonctions ------------------

def set_param_to_plot():
    for i, int in enumerate(param_ticked):
        param_to_plot[param_list[i]] = int.get()

def set_sample_to_plot():
    for i, int in enumerate(sample_ticked):
        sample_to_plot[sample_list[i]] = int.get()

def set_week_to_plot():
    for i, int in enumerate(week_ticked):
        week_to_plot[week_list[i]] = int.get()

def get_param_list(raw_data, line):
    global param_list
    global param_to_plot
    global param_col
    global param_unit
    param_list = []
    param_to_plot = {}
    param_col = {}
    param_unit = {}
    i=0

    p_list = raw_data.iloc[line-1:line+1]
    for parameter in p_list.values[0]:
        if isinstance(parameter, str):
            param_list.append(parameter)
            param_col[parameter] = i
            param_unit[parameter] = p_list.values[1][i]
            param_to_plot[parameter] = 0
        i+=1

def display_param(raw_data, line):
    global param_ticked
    param_ticked = []
    l = 0
    c = 0
    tk.Label(window, text = "Paramètres").place(x = 700, y = 20)
    get_param_list(raw_data, line)
    for param in param_list:
        param_ticked.append(tk.IntVar())
        r = tk.Checkbutton(window, text=param, variable=param_ticked[-1], command=set_param_to_plot)
        r.place(x=460+c, y= 63+l)
        c+=100
        if(c == 600):
            c=0
            l+=30

def get_sample_list(raw_data, col, first_sample, nb_sample):
    global sample_list
    global sample_to_plot
    sample_list = []
    sample_to_plot = {}
    first_sample_found = False
    n=0
    
    s_list = raw_data.iloc[0:, col:(col+1)]
    for s in s_list[col]:
        if first_sample_found == True:
            sample = ''.join(i for i in s if not i.isdigit())
            sample_list.append(sample)
            sample_to_plot[sample] = 0
            n+=1
            if n == nb_sample:
                break
            continue
        if isinstance(s, str):
            if s == first_sample:
                sample = ''.join(i for i in s if not i.isdigit())
                sample_list.append(sample)
                sample_to_plot[sample] = 0
                first_sample_found = True
                n+=1


def display_sample_list(raw_data, col, first_sample, nb_sample):
    global sample_ticked
    sample_ticked = []
    l=0
    c=0
    tk.Label(window, text = "Echantillons").place(x = 700, y = 180)

    if isinstance(col, str):
        col = ord(col) - ord("A")

    get_sample_list(raw_data, col, first_sample, nb_sample)
    
    for sample in sample_list:
        sample_ticked.append(tk.IntVar())
        r = tk.Checkbutton(window, text=sample, variable=sample_ticked[-1], command=set_sample_to_plot)
        r.select()
        r.place(x=460+c, y= 203+l)
        c+=100
        if(c == 600):
            c=0
            l+=30

    for int in sample_ticked:
        int.set(1)
    set_sample_to_plot()

def get_week_list(raw_data, col, first_sample):
    global week_list
    global week_lines
    global week_to_plot
    week_list = []
    week_lines = {}
    week_to_plot = {}
    i=0

    first_sample = ''.join(i for i in first_sample if not i.isdigit())
    
    s_list = raw_data.iloc[0:, col:(col+1)]
    for week in s_list[col]:
        if isinstance(week, str):
            if week[0] == first_sample:
                s = list(week)
                s[0] = "T"
                st = ''.join(s)
                week_list.append(st)
                week_to_plot[st] = 0
                week_lines[st] = i
        i+=1

def display_week_list(raw_data, col, first_sample):
    global week_ticked
    week_ticked = []
    l=0
    c=0
    tk.Label(window, text = "Semaines").place(x = 700, y = 280)

    if isinstance(col, str):
        col = ord(col) - ord("A")

    get_week_list(raw_data, col, first_sample)
    
    for week in week_list:
        week_ticked.append(tk.IntVar())
        r = tk.Checkbutton(window, text=week, variable=week_ticked[-1], command=set_week_to_plot)
        r.place(x=460+c, y= 303+l)
        c+=100
        if(c == 600):
            c=0
            l+=30

def extract_week_data(data_param, week):
    line = week_lines[week]

    week = data_param.iloc[line:line+12]
    week = week.rename(index={line: "A",line+1: "B",line+2: "C",line+3: "D",line+4: "E",line+5: "F",line+6: "G",line+7: "H",line+8: "I",line+9: "J",line+10: "K",line+11: "L"})
    return week.to_dict()

def plot_view(data, param, plot_id):
    abscisse = []
    sample_values = {}
    tmp_values = []

    plt.figure(plot_id)

    for week, has_to_be_plotted in week_to_plot.items():
        if has_to_be_plotted == 1:
            abscisse.append(week)

    for sample, has_to_be_plotted in sample_to_plot.items():
        if has_to_be_plotted == 1:
            for week, sample_measures in data.items():
                tmp_values.append(sample_measures[param][sample])
        
            sample_values[sample] = tmp_values.copy()
            tmp_values.clear()

    i = 0

    for sample, has_to_be_plotted in sample_to_plot.items():
        if has_to_be_plotted == 1:
            plt.plot(abscisse, sample_values[sample], label=sample, color=color_list[i], linewidth=2)
        i+=1
    
    plt.title("Evolution du "+param+" de "+abscisse[0]+" à "+abscisse[len(abscisse)-1])
    
    plt.legend(loc='upper right')
    plt.xlabel("Temps (mois)")
    plt.ylabel(param_unit[param])

def load_file():
    global raw_data
    filename = fd.askopenfilename(filetypes=(("xlsx files","*.xlsx"),("All files","*.*")))
    entry_path.insert(tk.END, filename)
    raw_data = pd.read_excel(filename, index_col=None, header=None)
    display_param(raw_data, int(entry_param.get()))
    display_sample_list(raw_data, entry_sample_col.get(), entry_sample_1.get(), int(entry_sample_nb.get()))
    display_week_list(raw_data, entry_sample_col.get(), entry_sample_1.get())

def go_graph():
    plot_id=0

    for param, has_to_be_plotted in param_to_plot.items():
        if has_to_be_plotted == 1:
            collect_data(param, plot_id)
            plot_id+=1

    plt.show()

def go_from_keyboard_event(event):
    go_graph()

def collect_data(param, plot_id):
    # Reducing excel sheet to only parameter's data
    data_param = raw_data.iloc[0:, param_col[param]:(param_col[param]+1)]

    # Renaming serie with friendly param name
    data_param = data_param.rename(columns={param_col[param]: param})

    # Getting each week param values in a dictionary
    data = {}
    for week, has_to_be_plotted in week_to_plot.items():
        if has_to_be_plotted == 1:
            data[week] = extract_week_data(data_param, week)

    # Affichage du graphe
    plot_view(data, param, plot_id)

if __name__ == '__main__':
    global window
    window = tk.Tk()
    window.title("Excetext_paramgraph")
    window.iconbitmap("./icon.ico")
    window.geometry("1100x700")
    window.resizable(0, 0)

    text_path = tk.Label(window, text = "Fichier excel")
    text_path.place(x = 20, y = 23)
    entry_path = tk.Entry(window, bd = 3, width=15)
    entry_path.place(x = 140, y = 20)
    button_path=tk.Button(window, text = "Chercher...", command=load_file, width=10)
    button_path.place(x = 300, y = 20)

    text_param = tk.Label(window, text = "Ligne des paramètres")
    text_param.place(x = 20, y = 63)
    entry_param = tk.Entry(window, bd = 3, width=2)
    entry_param.insert(tk.END, "7")
    entry_param.place(x = 255, y = 60)

    text_sample_col = tk.Label(window, text = "Colonne des échantillons")
    text_sample_col.place(x = 20, y = 103)
    entry_sample_col = tk.Entry(window, bd = 3, width=2)
    entry_sample_col.insert(tk.END, "C")
    entry_sample_col.place(x = 255, y = 100)

    text_sample_1 = tk.Label(window, text = "Premier échantillon")
    text_sample_1.place(x = 20, y = 143)
    entry_sample_1 = tk.Entry(window, bd = 3, width=2)
    entry_sample_1.insert(tk.END, "A0")
    entry_sample_1.place(x = 255, y = 140)

    text_sample_nb = tk.Label(window, text = "Nombre d'échantillons")
    text_sample_nb.place(x = 20, y = 183)
    entry_sample_nb = tk.Entry(window, bd = 3, width=2)
    entry_sample_nb.insert(tk.END, "12")
    entry_sample_nb.place(x = 255, y = 180)

    btn_go = tk.Button(window, text = "Go !", command=go_graph, width=20)
    btn_go.place(x = 700, y = 660)

    text_version = tk.Label(window, text = "Version 2021.11.12")
    text_version.place(x = 20, y = 670)

    window.bind('<Return>', go_from_keyboard_event)
    
    window.mainloop()

    sys.exit