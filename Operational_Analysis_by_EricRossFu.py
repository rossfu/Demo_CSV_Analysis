import os
import sys
import time
import subprocess
import tkinter as tk
from datetime import datetime
from tkinter import filedialog



# Initialize User Interface
GUI = tk.Tk()
GUI.title("")
GUI.geometry("600x450")
GUI.columnconfigure(0,weight=2)
GUI.columnconfigure(1,weight=1)
GUI.columnconfigure(2,weight=2)
GUI.configure(bg='white')
Header = tk.Label(GUI, text = '\n\nWelcome to your Operational Analysis Application.\n\n', font=('Arial', 14), bg='white')
Header.grid(column=1,row=0)
csv_file_list = []

    

# Analysis Function
def Analyze_Folder_Data(folder):

    # Install Pre-Requisites?
    try:
        import pandas as pd
        import numpy as np
        from iapws import IAPWS97
        import matplotlib.pyplot as plt
        import openpyxl
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        print('\nPre-requisites are imported.\n')
        
    except:
        install = tk.messagebox.askyesno('Confirm', 'Do you want to install: \n\npandas \nnumpy \nmatplotlib \nopenpyxl \niapws?')

        if install == True:
            subprocess.run(['pip', 'install', 'pandas'])
            subprocess.run(['pip', 'install', 'numpy'])
            subprocess.run(['pip', 'install', 'matplotlib'])
            subprocess.run(['pip', 'install', 'openpyxl'])
            subprocess.run(['pip', 'install', 'iapws'])

            import pandas as pd
            import numpy as np
            from iapws import IAPWS97
            import matplotlib.pyplot as plt
            import openpyxl
            from openpyxl import Workbook
            from openpyxl.utils.dataframe import dataframe_to_rows
            print('\nPre-requisites are imported.\nRunning Analysis...\n')
            
        else:
            tk.messagebox.showinfo('Notice', 'This application requires \n\npandas \nnumpy \nmatplotlib \nopenpyxl \niapws')
            #os._exit(0)

        
    # Grab every csv file
    for root,dirs,files in os.walk(folder):
        for names in files:
            if names.endswith('.csv'):
                csv_file_list.append(os.path.join(root,names))
                csv_file_list.sort(key=os.path.getmtime)

    print(str(len(csv_file_list)) + ' csv files detected.')
    print('\nAnalyzing the data in: ' + folder + '\n')
    

    # Append files together
    dataset = pd.read_csv(csv_file_list[0], sep = ',', header = 0, encoding= 'unicode_escape')     

    for file in csv_file_list[1:]:
        file_content = pd.read_csv(file, sep = ',', skiprows=0, encoding= 'unicode_escape')    
        dataset = pd.concat([dataset, file_content], ignore_index=True)
        


#########################################################################################################
#   Problem 1   (Data)
#########################################################################################################

    # Initialize Variables
    quarter = []
    q1 = ['01','02','03']
    q2 = ['04','05','06']
    q3 = ['07','08','09']
    q4 = ['10','11','12']

    temp = []
    temp2 = []
    press_mpa = []

    enthalpy_idx = []
    enthalpy_idx_val = 0
    enthalpy = []
    total_enthalpy = []
    enthalpy_quarter = []
    enthalpy_tix = [0]

    print(str(len(dataset)) + ' records detected.')

    for i in range(len(dataset)):

        #Gather 'Quarters'
        y = dataset['Timestamp'][i].split('-')[0]
        m = dataset['Timestamp'][i].split('-')[1]
        if m in q1:
            current_qtr = y + '_Q1'
            quarter.append(current_qtr)
        elif m in q2:
            current_qtr = y + '_Q2'
            quarter.append(current_qtr)
        elif m in q3:
            current_qtr = y + '_Q3'
            quarter.append(current_qtr)
        elif m in q4:
            current_qtr = y + '_Q4'
            quarter.append(current_qtr)
        else:
            print('quarter not identified')
            


        # Gather Temperature Ranges if online
        if dataset['Power (MW)'][i] > 30:
            
            if dataset['Temp (°F)'][i] <= 1000:             #cold
                temp.append('cold')
            elif 1000 < dataset['Temp (°F)'][i] <= 1050:    #norm
                temp.append('norm')
            elif 1050 < dataset['Temp (°F)'][i] <= 1100:    #hot
                temp.append('hot')
            elif 1100 < dataset['Temp (°F)'][i] <= 1150:    #real hot
                temp.append('very hot')
            else:
                temp.append('temperature outlier')

        else:
            temp.append('not online')
            
                

#########################################################################################################           #21872 - 21873
#   Problem 2   (Data)                                                                                              #32722 - 32914
#########################################################################################################           

        # Gather Temperature (Kelvin) Var
        current_temp_kelvin = ((dataset['Temp (°F)'][i] - 32) * (5/9)) + 273.15
##        temp2.append(current_temp_kelvin)

        # Gather Press_MPa Var
        current_press_mpa = (dataset['Press (psig)'][i] + 14.7) * 6894.76 / 1000000
##        press_mpa.append(current_press_mpa)
        

            
        # Does the Record meet Power and Temperature Reqs?
        # Yes
        if dataset['Power (MW)'][i] > 30 and dataset['PowerSwing (MW)'][i] <= 3 and dataset['Temp (°F)'][i] > 1000:


            # Calculate Enthalphy
            try:
                steam = (IAPWS97(P=current_press_mpa, T=current_temp_kelvin))
                current_enthalpy = steam.h * 0.4299226
                print('Analyzing record ' + str(i+1) + ' out of ' + str(len(dataset)) + '. Enthalpy is: ' + str(current_enthalpy))
            except:
                print('Enthalpy for Record # ' + str(i) + ' could NOT be calculated. Setting to zero.')
                current_enthalpy = 0

            # Chart 2 vars
            enthalpy.append(current_enthalpy)
            total_enthalpy.append(current_enthalpy)
            enthalpy_quarter.append(current_qtr)
            enthalpy_idx.append(enthalpy_idx_val)
            enthalpy_idx_val+=1

            try:
                if current_qtr != enthalpy_quarter[-2]:
                    enthalpy_tix.append(enthalpy_idx_val)
            except:
                pass

            
        # No
        else:
            print('Analyzing record ' + str(i+1) + ' out of ' + str(len(dataset)) + '. Power & Temperature Reqs not met.')
            total_enthalpy.append('Req not met')
    
###########################################
        
    # Add new variables to dataset
    dataset['Quarter'] = quarter
    dataset['Temp_Range'] = temp
##    dataset['Temp (K)'] = temp2
##    dataset['Press_MPa'] = press_mpa
##    dataset['Enthalpy'] = total_enthalpy
    

    # Write Report Into Excel
    wb = Workbook()
    ws = wb.active
    ws.title = 'Operational Report'



#########################################################################################################
# Chart 1 & Calculations              
#########################################################################################################
    unique_quarters = []
    number_of_hot_records = []
    number_of_very_hot_records = []
    
    for q in dataset['Quarter'].unique():                                                   # For each quarter

        unique_quarters.append(q)
        q_data = dataset[dataset['Quarter'] == q]
        
        if q_data[q_data['Temp_Range'] == 'hot'].shape[0] / 2 > 5:      
            number_of_hot_records.append(q_data[q_data['Temp_Range'] == 'hot'].shape[0] / 2)# Count hot records
        else:
            number_of_hot_records.append(0)                                                 # Unless less than 5 hours

            
        if q_data[q_data['Temp_Range'] == 'very hot'].shape[0] / 2 > 5:                    # Same for very hot records
            number_of_very_hot_records.append(q_data[q_data['Temp_Range'] == 'very hot'].shape[0] / 2)
        else:
            number_of_very_hot_records.append(0)


    # Plot    
    plt.figure(1)
    x = np.arange(len(dataset['Quarter'].unique()))
    width = 0.5
    rects1 = plt.bar(x - width/2, number_of_hot_records, width, label='1000-1050 deg F')
    rects2 = plt.bar(x + width/2, number_of_very_hot_records, width, label='1050-1100 deg F')

    # Add some Chart labels
    plt.xlabel('Date')
    plt.ylabel('Number of hours spent in Temperature Range')
    plt.title('Hours Spent in Hot Temperature Ranges By Quarter')
    plt.xticks(x, unique_quarters, rotation = 30)
    plt.legend()
    
    # Attach a text label above each bar
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            plt.annotate('{}'.format(height),
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 2),  # 3 points vertical offset
                        textcoords="offset points",
                        ha='center', va='bottom')
    autolabel(rects1)
    autolabel(rects2)
    

    # Save and Write Chart
    plt.savefig(selected_folder + '\\' + 'chart1.png')
    chart1 = openpyxl.drawing.image.Image(selected_folder + '\\' + 'chart1.png')
    chart1.anchor = 'A1'
    ws.add_image(chart1)



#########################################################################################################
# Chart 2 & Calculations              
#########################################################################################################


    fig2 = plt.figure(2)
    plot = plt.scatter(enthalpy_idx, enthalpy, s= .5, c = enthalpy, cmap='coolwarm')
    fig2.colorbar(plot)

    # Add some Chart labels
    plt.xlabel('Date')
    plt.ylabel('Enthalpy')
    plt.title('Enthalpy of Operating Equipment Over Time')
    tix = enthalpy_tix
    plt.xticks(tix, list(dataset['Quarter'].unique()), rotation = 30)
    
    # Save and Write Chart
    plt.savefig(selected_folder + '\\' + 'chart2.png')
    chart2 = openpyxl.drawing.image.Image(selected_folder + '\\' + 'chart2.png')
    chart2.anchor = 'L1'
    ws.add_image(chart2)

#########################################################################################################
    
    # Write Dataset / Calculations into Excel
    dataset_report = wb.create_sheet(title="Operational Dataset")
    for r in dataframe_to_rows(dataset, index=True, header=True):
        dataset_report.append(r)


    # Save / Open Report
    report_path = selected_folder + '\\' + 'ERIC ROSS FU_____Operational Report.xlsx'
    wb.save(report_path)
    print('\nGenerating Report at: ' + report_path)
    os.startfile(report_path)

    
    

def Select_Folder_Function():
    
    global selected_folder
    
    # Instructions and Folder Selection
    tk.messagebox.showinfo(title = None, message= 'Select the folder with your Operational Data')
    selected_folder = tk.filedialog.askdirectory(title='Select the folder with your Operational Data')

    # Proceed to analysis
    Analyze_Folder_Data(selected_folder)



# The Button   
folder_select_button = tk.Button(text='Perform Analysis', command=Select_Folder_Function, font=('Arial', 12), bg='white')
folder_select_button.grid(column=1,row=1)



