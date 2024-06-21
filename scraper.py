import tkinter as tk
from tkinter import *
import win32com.client as win32
import re, csv, os
from datetime import datetime
import pandas as pd
from prettytable import *
from time import sleep

def read_outlook_folder():

    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox_folder = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Find the PMS folder among the other ones under the Inbox
    pms_folder = None
    for folder in inbox_folder.Folders:
        if folder.Name == 'PMS':
            pms_folder = folder
            break 

    if pms_folder is not None:
        # Access the emails in the PMS folder
        emails = pms_folder.Items
        starter_list=[]   

        for email in emails:
            subject = email.Subject
            received_at = email.ReceivedTime.strftime("%d-%m-%Y %H:%M:%S")
               
            #Extract Server's name from mail title with RegEx
            srvname = re.search(r"\.(.*?):", subject)  #only match something between "." and ":"
            if srvname:
                srvname_ext = srvname.group(1)
                    
            #Extract Event (connect/disconnect) from mail title with RegEx   
            status = re.search(r"\((.*?)\)", subject) #only match something that is inside parentheses
            if status:
                status_ext=status.group(1)
            else:
                status_ext=""
                    
            try:   
            #Form the row
                row=[srvname_ext, status_ext, received_at]      

                #append the row to a list
                starter_list.append(row)

                sleep(0.01)

            except:
                print("\nInvalid E-mail format. Check PMS folder.")
    else:
        print(f"Folder 'PMS' not found.")
        
    # Release COM objects
    del outlook
    del namespace

    sort_list(starter_list)
    categorize()

def sort_list(input_list):

    # Read CSV data into a list of tuples
    data = input_list

    # Sort the data based on the datetime field
    sorted_list = sorted(data, key=lambda x: datetime.strptime(x[2], "%d-%m-%Y %H:%M:%S"))

    # Write the sorted data back to the CSV file - NEW LIST HERE
    with open("temp.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerows(sorted_list)

    header_ls = ['name', 'status', 'datetime']
    df = pd.read_csv("temp.csv", header=None)
    df.to_csv("temp.csv", header=header_ls, index=False) 

def categorize():

    data = pd.read_csv("temp.csv")
    nameslist = data['name'].tolist()
    statlist = data['status'].tolist()
    datelist = data['datetime'].tolist()

    no_tuples = []

    for i in nameslist:
        if i not in no_tuples:
            no_tuples.append(i)

    up=0
    down=0
    unkn=0
    n=0

    tableOutput = PrettyTable(["No.", "Server", "Latest Status", "Date & Time"])
    tableOutput.set_style(SINGLE_BORDER)

    for i in no_tuples:
            
        n+=1

        for (j ,k, l) in zip(nameslist, statlist, datelist):
            if i==j:
                status=k
                event=l

        match status:

            case 'connect':
                up+=1
                printstat='OK'

            case 'disconnect':
                down+=1
                printstat='Disconnected'

            case _:
                unkn+=1
                printstat='Unknown/Recovering'

        tableOutput.add_row([n, i, printstat, event])
        sleep(0.1) #do not remove else non-compatible emails are detected, even if they don't exist

    txtbox.insert(END, tableOutput)
    #print('\n'+str(up)+' connected, '+str(down)+' disconnected, '+str(unkn)+' in unknown status\n')

    os.remove("temp.csv")

w=tk.Tk()
w.iconbitmap("icon.ico")
w.title("PMS Scraper")
w.resizable(width=False, height=False)

txtbox = Text(w, height=15, width=62)
connected_lbl = tk.Label(textvariable="Connected:").pack()
button = tk.Button(text='Start', width=25, command=read_outlook_folder)

button.pack(pady=15)
txtbox.pack(pady=10)

w.mainloop()
