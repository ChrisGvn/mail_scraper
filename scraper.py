import win32com.client as win32
import re, csv
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
        #f = open("status.csv", "w",  newline='') #new list here
        #writer=csv.writer(f)
        
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
    #with open("status.csv", "r", newline="") as file:
        #reader = csv.reader(file)
        
        #for row in reader:
            #data.append(row)

    # Sort the data based on the datetime field
    sorted_data = sorted(data, key=lambda x: datetime.strptime(x[2], "%d-%m-%Y %H:%M:%S"))

    # Write the sorted data back to the CSV file
    with open("status.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerows(sorted_data)

    print(f"CSV file has been sorted based on the datetime field.\n")

def categorize():

    header_ls = ['name', 'status', 'datetime']
    df = pd.read_csv("status.csv", header=None)
    df.to_csv("status.csv", header=header_ls, index=False)

    data = pd.read_csv("status.csv")
    nameslist = data['name'].tolist()

    newlist = []
    for i in nameslist:
        if i in newlist:
            print(i+' is already in')
        else:
            newlist.append(i)
            print(i+' added to list')

    ##junk = pd.read_csv("status_sorted.csv")
    ##namejunk = junk['name'].tolist()
    ##statjunk = junk['status'].tolist()

    for i in newlist:

        status='disconnected'
        #for j in junk:
    #       if name -> match with something
    #           search if connect or disconnect
    #           
    #      

    print(newlist)

# Specify the custom folder name
folder_name = "PMS"

# Call function
read_outlook_folder(folder_name)
#sort_list()
categorize()
