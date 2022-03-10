# coding: utf-8
#!/usr/bin/python

# ------------------ Librairies ------------------
import sys
import os
import pandas as pd
import numpy as np
import openpyxl as px
import matplotlib.pyplot as plt
import tkinter as tk
import tkinter.filedialog as fd

import codecs
if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

color_list = ["brown","royalblue","orchid","aqua","red","steelblue","lime","orange","darkseagreen","pink","gold","darkgreen","rosybrown",]

# ------------------ Fonctions ------------------

# Gives the week list to plot from user input weeks
def get_weeks_to_plot():
    weeks_to_display = []

    for i in range(len(week_list)):
        if weeks_check_button[i].get() == 1:
            weeks_to_display.append(week_list[i])
    
    return weeks_to_display

def errorfill(x, y, yerr, color=None, alpha_fill=0.2, ax=None):
    ax = ax if ax is not None else plt.gca()
    if color is None:
        color = ax._get_lines.color_cycle.next()
    if np.isscalar(yerr) or len(yerr) == len(y):
        ymin = np.array(y) - np.array(yerr)
        ymax = np.array(y) + np.array(yerr)
    elif len(yerr) == 2:
        ymin, ymax = yerr
    ax.plot(x, y, color=color)
    ax.fill_between(x, ymax, ymin, color=color, alpha=alpha_fill)

# Gives row number in excel file from user input parameter
def get_param_row(parameter):
    p_list = raw_data.iloc[6:7].values[0]
    row = 0
    for p in p_list:
        if p == parameter:
            return row
        row+=1

# Gives week start line in excel file from user input parameter
def get_week_start_line(week):
    return week_start_line[week]

# Reads user input parameter units in file
def get_param_unit(param_data):
    unit_cell = param_data.iloc[7:8].values[0]
    if isinstance(unit_cell[0], str) == False:
        return "Sans unité"
    return unit_cell[0]

# Extract data from a specified week
def extract_week(param_data, week):
    line = get_week_start_line(week)

    week = param_data.iloc[line:line+len(sample_list)]
    for i in range(len(sample_list)):
        week = week.rename(index={line+i:sample_list[i]})

    # week = week.rename(index={line: "A",line+1: "B",line+2: "C",line+3: "D",line+4: "E",line+5: "F",line+6: "G",line+7: "H",line+8: "I",line+9: "J",line+10: "K",line+11: "L"})
    return week.to_dict()

def plot_view(data, error, param, displayed_week_list, unit):
    x_legend = []
    displayed_values = {}
    displayed_errors = {}
    sample_values = []
    error_values = []

    Title = param+" "+displayed_week_list[0]+" - "+displayed_week_list[len(displayed_week_list)-1]

    plt.figure(Title)

    for key, value in data.items():
        x_legend.append(week_days_from_start[key]) # Getting T0, T1, T1+1 etc...

    for i in range(len(sample_list)):
        if sample_check_button[i].get() == 1:
            for key, value in data.items():
                if (isinstance(value[param][sample_list[i]], float) == False) and (isinstance(value[param][sample_list[i]], int) == False):
                    try:
                        sample_values.append(float(value[param][sample_list[i]][1:]))
                    except ValueError:
                        sample_values.append(value[param][sample_list[i]])
                else:
                    sample_values.append(value[param][sample_list[i]]) # Getting A value for T0, then B value for T0, etc...
            for key, value in error.items():
                try:
                    error_values.append(value["Error_"+param][sample_list[i]]/2) # Getting A error value for T0, then B error value for T0, etc...
                except TypeError:
                    error_values.append(value["Error_"+param][sample_list[i]])
            
            displayed_values[sample_list[i]] = sample_values.copy()
            displayed_errors[sample_list[i]] = error_values.copy()
            sample_values.clear()
            error_values.clear()
        
    for i in range(len(sample_list)):
        if sample_check_button[i].get() == 1:
            plt.plot(x_legend, displayed_values[sample_list[i]], label=sample_list[i], color=color_list[i])
            if error_visibility.get() == 1 and param_has_error_data and (isinstance(displayed_values[sample_list[i]][0], str)==False) and (isinstance(displayed_errors[sample_list[i]][0], int) or isinstance(displayed_errors[sample_list[i]][0], float)):
                errorfill(x_legend, displayed_values[sample_list[i]],  displayed_errors[sample_list[i]], color_list[i], alpha_error_value)

    plt.title("Evolution du "+param+" de "+displayed_week_list[0]+" à "+displayed_week_list[len(displayed_week_list)-1])
    if grid_visibility.get() == 1:
        plt.grid(visible=True, which='major', color='#888888', linestyle='-', alpha=0.5)
        plt.minorticks_on()
        plt.grid(visible=True, which='minor', color='#999999', linestyle='-', alpha=0.2)

    plt.legend(loc='upper right', bbox_to_anchor=(1.1, 1.0))
    plt.xlabel("Temps (jours)")
    plt.ylabel(unit)

def compute(param):
    global param_has_error_data
    param_has_error_data = False

    # Getting column num for user input parameter
    row = get_param_row(param)

    # Reducing excel sheet to only parameter's data
    param_data = raw_data.iloc[0:, row:(row+1)]
    param_data_error = raw_data.iloc[0:, (row+1):(row+2)]

    # Renaming serie with friendly param name
    # param_data = param_data.rename(columns={row: param})
    # param_data_error = param_data_error.rename(columns={row: "error_"+param})
    param_data.columns = [param]
    
    if not param_data_error.empty:
     # insert code here
        param_has_error_data = True
        param_data_error.columns = ["Error_"+param]
    

    # Getting unit
    unit = get_param_unit(param_data)

    # Getting weeks list needed
    displayed_week_list = get_weeks_to_plot()

    # Getting each week in a dictionary
    data = {}
    error = {}
    for week in displayed_week_list:
        data[week] = extract_week(param_data, week)
        if param_has_error_data:
            error[week] = extract_week(param_data_error, week)

    # Affichage du graphe
    plot_view(data, error, param, displayed_week_list, unit)


# ------------------ Affichage ------------------

def browsefunc():
        filename = fd.askopenfilename(filetypes=(("xlsx files","*.xlsx"),("All files","*.*")))
        e1.insert(tk.END, filename)

def compute_data_from_event(event):
    compute_data()

def compute_data():
    # Traite la valuer de transparence de la barre d'erreur
    global alpha_error_value
    alpha_error_value = e_alpha_error.get()
    try: 
        alpha_error_value=int(alpha_error_value)/100
    except ValueError:
        alpha_error_value = 0.2
    if alpha_error_value < 0.0 or alpha_error_value > 1.0:
        alpha_error_value = 0.2

    # Vérifie qu'il y a bien au moins 1 echantillon de selectioné
    selected_samples = []
    for i in range(len(sample_list)):
        if sample_check_button[i].get() == 1:
            selected_samples.append(sample_list[i])
    
    if len(selected_samples) < 1:
        return


    # Vérifie qu'il y a bien au moins 2 semaines de selectionées
    selected_weeks = []
    for i in range(len(week_list)):
        if weeks_check_button[i].get() == 1:
            selected_weeks.append(week_list[i])
    
    if len(selected_weeks) < 2:
        return


     # Affiche le graph pour un paramètre donné
    for i in range(len(param_check_button)):
        if param_check_button[i].get() == 1:
            compute(parameter_list[i])

    plt.show()

def read_excel_data_from_event(event):
    read_excel_data()

def read_excel_data():
    global path
    path = e1.get()
    load_excel_data()
    hide_file_browse()
    display_settings()

def load_excel_data():
    global parameter_list
    global sample_list
    global raw_data
    global week_start_line
    global week_list
    global week_days_from_start

    # Read excel file
    raw_data = pd.read_excel(path, sheet_name="Données",index_col=None, header=None)

    parameter_list = raw_data.iloc[6:7].values[0]
    parameter_list = [x for x in parameter_list if pd.isnull(x) == False]

    week_list = []
    week_days_from_start = {}
    week_start_line = {}
    sample_list = []
    first_sample_found = False
    last_sample_found = False

    week_row = raw_data.iloc[:,2]
    days_row = raw_data.iloc[:,1]
    for index in range(len(week_row)):

        if first_sample_found == True:
            if last_sample_found == False:
                if str(week_row[index])[:1] == "A":
                    last_sample_found = True
                else:
                    sample_list.append(week_row[index][0])

            if str(week_row[index])[:1] == "A":
                week_name = 'T'+ week_row[index][1:]
                week_list.append(week_name)
                week_start_line[week_name] = index
                try:
                    week_days_from_start[week_name] = int(days_row[index])
                except:
                    week_days_from_start[week_name] = days_row[index]

        elif str(week_row[index])[:1]== "A":
            first_sample_found = True
            sample_list.append(week_row[index][0])
            week_name = 'T'+ week_row[index][1:]
            week_list.append(week_name)
            week_start_line[week_name] = index
            n = days_row[index]
            week_days_from_start[week_name] = int(days_row[index])


def hide_file_browse():
    l1.place_forget()
    e1.place_forget()
    b1.place_forget()
    btn.place_forget()

def display_settings():
    global sample_check_button
    global param_check_button
    global weeks_check_button
    global error_visibility
    global e_alpha_error
    global grid_visibility

    upper_text_px_height = 0

    l5 = tk.Label(window, text = "Echantillons")
    l5.place(x = 10, y = upper_text_px_height+20)

    sample_check_button = [0]*len(sample_list)
    i = 0
    x_offset=0
    y_offset=0
    for sample in sample_list:
        sample_check_button[i] = tk.IntVar()
        e = tk.Checkbutton(window, variable=sample_check_button[i], text=sample, onvalue=1, offvalue=0, fg=color_list[i])
        e.place(x = 100 + (40*x_offset), y = upper_text_px_height + 20 + (20*y_offset))
        e.select()
        i+=1
        x_offset+=1
        if (i%12 == 0):
            y_offset+=1
            x_offset=0

    upper_text_px_height += (y_offset+1) * 20 

    l6 = tk.Label(window, text = "Paramètres")
    l6.place(x = 10, y = upper_text_px_height+20)

    param_check_button = [0]*len(parameter_list)
    i = 0
    x_offset=0
    y_offset=0
    for param in parameter_list:
        param_check_button[i] = tk.IntVar()
        e = tk.Checkbutton(window, variable=param_check_button[i], text=param, onvalue=1, offvalue=0)
        e.place(x = 100 + (90*x_offset), y = upper_text_px_height + 20 + (20*y_offset))
        i+=1
        x_offset+=1
        if (i%8 == 0):
            y_offset+=1
            x_offset=0

    upper_text_px_height += 20 + (y_offset+1) * 20 

    
    l7 = tk.Label(window, text = "Période")
    l7.place(x = 10, y = upper_text_px_height+20)

    weeks_check_button = [0]*len(week_list)
    i = 0
    x_offset=0
    y_offset=0
    c = 'T0'
    for week in week_list:

        if (c != week[:2]):
            c = week[:2]
            y_offset+=1
            x_offset=0

        weeks_check_button[i] = tk.IntVar()
        e = tk.Checkbutton(window, variable=weeks_check_button[i], text=week+" (j"+str(week_days_from_start[week])+")", onvalue=1, offvalue=0)
        e.place(x = 100 + (100*x_offset), y = upper_text_px_height + 20 + (20*y_offset))
        i+=1
        x_offset+=1

    upper_text_px_height += 40 + (y_offset+1) * 20

    l8 = tk.Label(window, text = "Barre d'erreur")
    l8.place(x = 10, y = upper_text_px_height+20)
    error_visibility = tk.IntVar()
    cb_error_visibility = tk.Checkbutton(window, variable=error_visibility, onvalue=1, offvalue=0)
    cb_error_visibility.place(x = 100, y = upper_text_px_height + 20)
    cb_error_visibility.select()
    
    l_alpha_error = tk.Label(window, text = "Transparence (%)")
    l_alpha_error.place(x = 150, y = upper_text_px_height+20)
    e_alpha_error = tk.Entry(window, width=3)
    e_alpha_error.place(x = 250, y = upper_text_px_height + 20)
    e_alpha_error.insert(-1, '10')

    upper_text_px_height += 40

    l10 = tk.Label(window, text = "Quadrillage")
    l10.place(x = 10, y = upper_text_px_height+20)
    grid_visibility = tk.IntVar()
    cb_grid_visibility = tk.Checkbutton(window, variable=grid_visibility, onvalue=1, offvalue=0)
    cb_grid_visibility.place(x = 100, y = upper_text_px_height + 20)
    cb_grid_visibility.select()

    upper_text_px_height += 40

    l4.place(x = 10, y = upper_text_px_height + 40)
    btn = tk.Button(window, text = "Dessiner", command=compute_data, width=20)
    btn.place(x = 370, y = upper_text_px_height + 30)
    window.bind('<Return>', compute_data_from_event)
    window.geometry("1160x"+str(upper_text_px_height+80))


if __name__ == '__main__':
    
    window = tk.Tk()
    window.title("Excel2graph")
    if "nt" == os.name:
        window.iconbitmap('./icv.ico')
    else:
        window.iconbitmap('@icv.xbm')
    window.geometry("430x100")
    window.resizable(0, 0)

    l1 = tk.Label(window, text = "Fichier excel")
    l1.place(x = 10, y = 20)
    e1 = tk.Entry(window, bd = 5)
    e1.place(x = 140, y = 20)
    b1=tk.Button(window, text = "Chercher...", command=browsefunc)
    b1.place(x = 320, y = 20)    
    btn = tk.Button(window, text = "Ouvrir", command=read_excel_data, width=20)
    btn.place(x = 220, y = 60)
    l4 = tk.Label(window, text = "Version 2022.03.08-22h31")
    l4.place(x = 10, y = 70)

    window.bind('<Return>', read_excel_data_from_event)
    
    window.mainloop()

    sys.exit