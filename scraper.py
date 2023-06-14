import win32com.client as win32
import re, csv, itertools
from datetime import datetime
import pandas as pd

def read_outlook_folder(folder_name):
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox_folder = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Find the PMS folder under the Inbox
    pms_folder = None
    for folder in inbox_folder.Folders:
        if folder.Name == folder_name:
            pms_folder = folder
            break 

    if pms_folder is not None:
        # Access the emails in the PMS folder
        emails = pms_folder.Items
        
        #Email counter set to 0
        cn=0

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

            #Form the row
            row=[srvname_ext, status_ext, received_at]
                
            #append the row to a list
            starter_list.append(row)

            #Email counter +1
            cn+=1
    else:
        print(f"Folder '{folder_name}' not found.")
        
    # Release COM objects
    del outlook
    del namespace
    
    print("\nTotal of "+str(cn)+ " messages in the folder.")
    sort_list(starter_list)

def sort_list(input_list):

    # Read CSV data into a list of tuples
    data = input_list

    # Sort the data based on the datetime field
    sorted_data = sorted(data, key=lambda x: datetime.strptime(x[2], "%d-%m-%Y %H:%M:%S"))

    # Write the sorted data back to the CSV file
    with open("status.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerows(sorted_data)

    header_ls = ['name', 'status', 'datetime']
    df = pd.read_csv("status.csv", header=None)
    df.to_csv("status.csv", header=header_ls, index=False) 

    print(f"CSV file has been sorted based on the datetime field.\n")

def categorize():

    data = pd.read_csv("status.csv")
    nameslist = data['name'].tolist()
    statlist = data['status'].tolist()

    no_tuples = []
    for i in nameslist:
        if i not in no_tuples:
            no_tuples.append(i)

    for i in no_tuples:
        for (j ,k) in zip(nameslist, statlist):
            if i==j:
                status=k
        
        print(i+' '+status)
        
# Call functions
read_outlook_folder('PMS')
categorize()

input('\nPress any key to close...')
