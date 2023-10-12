import pandas as pd 
import os as os
import datetime as dt
import time

excel_file_master = r"~\Desktop\EGR KDL Master 3D Printing List.xlsx"

df = pd.read_excel(excel_file_master)

#making sure python understands this excel file is a data frame 
df = pd.DataFrame(df)

#Code trying to simplify the datetime variable to make it more readable to staff 
df['completion_time'] = df['completion_time'].dt.strftime("%b %d, %Y")

#Filtering the data frame by prints that have been completed and patrons contacted but the print has not been picked up yet and is not already marked as expired 
df = df[(df["print_completed"] == "X") & (df["patron_contacted"] == "X") & (df["picked_up"].isna()) & (df["print_expired"].isna())]


#Selecting specific varaibles to be displayed in the excel file so that it's for staff to see when looking at expired prints page  

#Select by selecting
#df = df[["ID" ,"completion_time", "name", "email", "phone_number", "print", "color"]]

#Select by dropping 
#Drop specific variables from our data frame
df = df.drop(columns= ["print_started", "invalid_email", "sheet_print", "failed"])


#Next we filter the contacted date so that three weeks has passed since the print has been finished but not picked up yet 

# calculate the date three weeks ago
three_weeks_ago = dt.datetime.now() - dt.timedelta(days=21)

# This will give you all the rows where 'contacted_date' is more than three weeks old.
df_filter = df[df['contacted_date'] < three_weeks_ago]

#Using the .loc method to select the rows where the patron has been contacted has been contacted and putting an "X" in the print_expired column
df_filter.loc[df_filter['patron_contacted'] == 'X', 'print_expired'] = 'X'   


#Showing the user the first 10 observations of the Expired 3D Prints List for verificiation 
import sys

print("\nBelow are the first 10 observations of the 'Expired 3D Prints List' DataFrame:\n")
print(df_filter.head(10))

while True:
    answer = input("\nSeeing the above DataFrame, would you like to continue with the program to save the above DataFrame locally as an excel file? (y/n) ")
    if answer == "y":
        print("Continuing with the program!") 
        break
    elif answer == "n":
        print("The program will now cease running.")
        sys.exit()
    else:
        print("Invalid input. Please enter 'y' to continue or 'n' to exit.")

# Continue running the program here


#The following code takes the df_filter dataframe that contains all the expired prints for people who have been completed longer than two weeks ago and saves it locally as an excel file

#Gets today's datetime value
today = dt.datetime.now().strftime("%b %d") #The format codes %b and %d result in "Feb 20"

string = "Updated_expired_Var"  #String we are searching all files in the desktop for  

#File name changes based on the day that the program is run 
file_name = f"Updated_expired_Var {today}.xlsx"


#Recreating the folder to save excel files if it doesn't already exist, os.path.expanduser is used because os.makedirs doesn't like the ~
home_dir = os.path.expanduser("~")
#stating the directory of where to save our new expired prints excel 
directory_save = os.path.join(home_dir, "Desktop", "3D Printing output files")
os.makedirs(directory_save, exist_ok=True)


#Specific path of where I want to save the new excel doc being created 
file_path_upadte = directory_save + "\\" + file_name

#Creating the second excel file that staff wil use to easily contact the affected people  
file_print = f"Expired 3D prints week of {today}.xlsx"

file_path_print = directory_save + "\\" + file_print

#loops through all the files in the directory to search for an older version of this document (file_name) (if it exists) then update it! 
for filename in os.listdir(directory_save): 
    if string in filename:
        file_pathy = os.path.join(directory_save, filename)
        answer = input(f"\nOld file exists, \'{filename}\', would you like to delete the this old one and save the updated versions? (y/n) ")

        if answer == 'y':
            os.remove(file_pathy)
            print(f"\nOld file: \'{filename}\' has been deleted\n")

        else:
            print(f"\nThe existing files will not be overwritten.")
            print("This program will now cease running.")
            sys.exit()

#Creates the new excel file if one is not already there 

#Code to create a new excel doc marking the patrons who have been contacted via email
master_df = pd.read_excel(excel_file_master) #loads in master excel doc from onedrive

master_df.update(df_filter) #updates master_df doc with the updated patron_contacted varaible from df_filter (AKA tallying patrons who need to be recontacted)

master_df.to_excel(file_path_upadte , index=False) #Saves the updated master_df to onedrive to be used to update the master excel doc (that the form feeds into)

#After the emails are sent we need to update the patron_contacted column so this will create another excel I'll use to update the master doc 
print(f"Use \'{file_name}\' excel sheet to update the 'print_expired' varaible in the Master 3D Printing doc!!!\n")

df_filter.to_excel(file_path_print , index=False)

print(f"Use \'{file_print}\' excel sheet to identify and remove the 3D prints that have been sitting for 3 weeks after the patron was initally contacted")


#Another prompt now that the excel files are saved locally seeing if we want to email it out to 3D print staff  
user_input = input(f"\nNow that \'{file_name}\' and \'{file_print}\' are saved locally, would you like to email them to 3D print staff? (y/n) ")

if user_input.lower() == "y":
    print("Prepping email to print to staff!\n")
    time.sleep(1)

else:
    print("With the file just saved locally, make sure to print it out!")
    time.sleep(4)
    sys.exit()

#Code to send an email with multiple attachments to multiple email addresses at once  

from email.message import EmailMessage #Library that formats code and strings as an email 
import ssl #adds security to our email 
import smtplib #library that actually sends the email 
import time  #used to get time.sleep function pause the program as to not send emails too fast and get blocked by gmail

 
#List of emails for EGR KDL 3D print team 
KDL_emails = ["cdelongchamp@kdl.org", "HGoulet@kdl.org", "AVuong@kdl.org", "RIacovone@kdl.org", "PLu@kdl.org", "JSavage-Dura@kdl.org",  "trhoades@kdl.org", "hmathews@kdl.org"]

email_sender = "egr3dprinting@gmail.com"
email_pass = "vpbpggszhbhnklkz"

subject = f"Expired Prints List for Week of {today}"

body = "Hello 3D Print Team," + "\n\n" + "The attached excel file contains the list of expired prints we can donate or discard. " + \
    "\n\n" + "Patrons who were contacted more than 3 weeks ago about their print and have not yet picked it up are on this list.\n" 

# list of file paths to attach to the email 
file_paths = [file_path_upadte, file_path_print]

#matching up the parameters above for our email to the format python expects them 
em = EmailMessage()
em['From'] = email_sender
em['To'] = KDL_emails
em['Subject'] = subject
em.set_content(body)

# Attaching the .xlsx file to the email
for file_path in file_paths:
    # get the file name from the file path
    file_name = os.path.basename(file_path)
    with open(file_path, "rb") as f:
        file_data = f.read() #File needs to be read in as binary code 
        em.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename= file_name)

#adding a layer of security to our email 
context = ssl.create_default_context()

# Sending an email to all KDL 3D print staff all at once
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_pass)
    smtp.sendmail(email_sender, KDL_emails, em.as_string())

print(f"\n\'{file_name}\' and {file_print} has been successfully emailed to EGR KDL 3D printing team!")