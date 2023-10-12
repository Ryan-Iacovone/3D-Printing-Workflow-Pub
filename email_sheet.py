##Code to send an updated version of the excel sheet to 3D print team every week

#Reading in the updated data from the microsoft form using onedrive 

import pandas as pd 
import os
import datetime
import time

excel_file_master = r"~\Desktop\EGR KDL Master 3D Printing List.xlsx"

#Reading in the updated data from the microsoft form using an excel doc stored on onedrive 
df = pd.read_excel(excel_file_master) 

#making sure python understands this excel file is a data frame 
df = pd.DataFrame(df)

#Filtering the data frame by patrons who's print has not ealready been completed and who's print has has not already gotten on the physical sheet  
df = df[(df["print_completed"] != "X") & (df["sheet_print"] != "X") ]

# format the 'number' column (NOT WORKING)
#df['card_number'] = df['card_number'].apply(lambda x: format(x, '.0f'))


    #Select by dropping 
#Drop specific variables from our data frame
df_selected = df.drop(columns= ["patron_contacted", "contacted_date", "invalid_email", "print_expired", "sheet_print", "recontacted?"])


#Sorting the base data frame primarily by color and then by date for easy read ability 

df = pd.DataFrame(df)
df_selected_sorted = df_selected.sort_values(by=['color', 'completion_time'],  ascending= [False, True])


#simplifying the datetime variable completion_time to make it more readable to staff 

df_selected_sorted['completion_time'] = pd.to_datetime(df_selected_sorted['completion_time'])
df_selected_sorted['completion_time'] = df_selected_sorted['completion_time'].dt.strftime("%b %d, %Y") #puts date in the format: "Feb 18, 2023"


#Initiating a prompt in the program to see if the user wants to continue with creating an excel file 
import sys

print("\nBelow are the first 10 observations of the '3D Prints to be Completed' DataFrame:\n")
print(df_selected_sorted.head(10))

while True:
    answer = input("\nSeeing the above DataFrame, would you like to make it an excel file and save it locally? (y/n) ")
    if answer == "y":
        break
    elif answer == "n":
        print("The program will now cease running.")
        time.sleep(3)
        sys.exit()
    else:
        print("Invalid input. Please enter 'y' to continue or 'n' to exit.")


#The following code saves the df_selected_sorted data frame I've edited and saves it locally as an excel doc 

today = datetime.datetime.now().strftime("%b %d") #gets the current datatime at the moment the code is run, formats it with strftime as 'Feb 20'

#Save excel file with name based on day program is run
file_name = f"3D print jobs for week of {today}.xlsx"

#Recreating the folder to save excel files if it doesn't already exist, os.path.expanduser is used because os.makedirs doesn't like the ~
home_dir = os.path.expanduser("~")
directory_save = os.path.join(home_dir, "Desktop", "3D Printing output files")
os.makedirs(directory_save, exist_ok=True)

#Creating the complete path (location) where I want to save the new excel file
file_path = directory_save + "\\" + file_name

#Code to see if the file_path already exists then if it does asks user if they want to overwrite it 
if os.path.exists(file_path):
    answer = input(f"\n{file_name} already exists, would you like to overwrite it with the new file? (y/n) ")

    if answer == 'y': 
        os.remove(file_path)
        print(f"\n\"{file_name}\" has successfully been saved\n")

    else:
        print(f"Old {file_name} file will not be overwritten")
        print("This program will now cease running...")
        time.sleep(3)
        sys.exit()

else:
    print(f"\n\"{file_name}\" has successfully been saved")

#Code to saved the excel file to the file_path specified above 
df_selected_sorted.to_excel(file_path, index=False) 


#Making a separate excel sheet marking which prints have been added to the physical printing sheet 

directory = r"~\Desktop\3D Printing output files" #Specifies the directory we're going to search for files and create the new excel file noting patrons contacted  
today = datetime.datetime.now().strftime("%b %d") #Gets today's today in the format of "Feb 18"
file_name2 = f"Updated sheet_print {today}.xlsx" #Creating a new excel file with the updated date in the name 
file_path2 = os.path.join(directory, file_name2) #Combining the directory and file name to create the file path to more easily write this file location later on 

#we need to read in the orginal file again to a dataframe to get the structure of it making copy and pasting the inputs back in way easier 
#Code to create a new excel doc marking the patrons who have been contacted via email
master_df = pd.read_excel(excel_file_master) #loads in master excel doc from onedrive

#Getting the list of varaibles that are on our 3D print list we want to keep track of 
identification = df_selected_sorted["ID"].tolist()

#Takes all the IDs of the patrons who are being put on the 3D print list for the week and puts an X in the sheet_print column noting that they've been printed
# master_df.loc uses the pandas locator method to locate the row(s) where the "ID" column equals the current id in the list we're itterating over. With those located rows, 
# it adds an "X" to the "sheet_print" column 
for id in identification:
   master_df.loc[master_df["ID"] == id, "sheet_print"] = "X" 

master_df.to_excel(file_path2 , index=False) #Saves the updated master_df to onedrive to be used to update the master excel doc (that the form feeds into)

print(f"\nAlso, new file: \'{file_name2}\' has successfully been created!\n")

print(f"Now, use \'{file_name2}\' document to update the varaible 'sheet_print' in the Master 3D Printing doc!!!\n")

time.sleep(4) 


#Another prompt now that the excel file is saved locally seeing if we want to email it out to 3D print staff  

user_input = input(f"Now that {file_name} is saved locally, would you like to email it out to 3D print staff? (y/n) ")

if user_input.lower() == "y":
    print("Prepping email to print to staff!\n")
else:
    print("With the file just saved locally, make sure to print it out!")
    time.sleep(4)
    sys.exit()


#Code to generate a random quote from the API link to add at the end of the weekly email

import requests
import json

def get_random_quote():
  url = "http://api.quotable.io/random"
  response = requests.get(url)
  if response.status_code == 200:
    quote = response.json()
    return quote["content"] + " - " + quote["author"]
  else:
    return "Error: Could not retrieve quote."

quote_of_the_week = 'Quote of the week: "' + get_random_quote() + '"'

quote_of_the_week


#Code to send an email to multiple email addresses at once with an excel attachment   

from email.message import EmailMessage #Library that formats code and strings as an email 
import ssl #adds security to our email 
import smtplib #library that actually sends the email 
import time  #used to get time.sleep function pause the program as to not send emails too fast and get blocked by gmail


#List of emails for EGR KDL 3D print team 
KDL_emails = ["cdelongchamp@kdl.org", "HGoulet@kdl.org", "AVuong@kdl.org", "RIacovone@kdl.org", "PLu@kdl.org", "JSavage-Dura@kdl.org",  "trhoades@kdl.org", "hmathews@kdl.org"]

email_sender = "egr3dprinting@gmail.com"
email_pass = "vpbpggszhbhnklkz"

subject = f"3D Print List for Week of {today}"

body = "Hello 3D Print Team," + "\n\n" + "Please print the attatched excel file containing the 3D prints to be completed for this week." + "\n\n" + "Printing instructions: " + \
    "Expand each of the columns so that each column name and their associated observations are readable. To make this easier highlight all of the column names and left align them." + \
    " Next, under page layout, set the top, bottom, left, and right margins close to 0. Set the page orientation to 'landscape'. Then finally, click the print gridlines option. " + "\n\n" + \
    f"{quote_of_the_week}" + "\n\n" 

#defining the path of the excel attachment to be sent 
new_week_excel = file_path

# Attaching the .xlsx file to the email
with open(new_week_excel, "rb") as f:
    file_data = f.read() #File needs to be read in as binary code 

#adding a layer of security to our email 
context = ssl.create_default_context()

#matching up the parameters above for our email to the format python expects them 
em = EmailMessage()
em['From'] = email_sender
em['To'] = KDL_emails
em['Subject'] = subject
em.set_content(body)
em.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename= file_name)

# Sending an email to all KDL 3D print staff all at once
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_pass)
    smtp.sendmail(email_sender, KDL_emails, em.as_string())

print("\nEmail successfully sent to EGR KDL printing team!")
time.sleep(3)