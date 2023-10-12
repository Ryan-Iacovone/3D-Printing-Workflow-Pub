import pandas as pd 

#Importing in master 3D printing data set from onedrive to get updated form data 
excel_file_master = r"~\Desktop\EGR KDL Master 3D Printing List.xlsx"

df = pd.read_excel(excel_file_master)

#filtering data by the people who's print is completed, they have not been contacted yet, they do not have an invalid email address, and there print has not been picked up yet
df_filter = df[(df["print_completed"] == "X") & (df["patron_contacted"].isna()) & (df["invalid_email"].isna()) & (df["picked_up"].isna())] #can do is.na or set patron_contacted = "X"


#Following code applies regex to a list of email addresses we gathered above then sorts by validness 
import re

#Gathering the list of emails from the data set who's prints have been completed (but not yet contacted) so that they can now be emailed 
emails = df_filter["email"].tolist()

#Creates 2 empty lists where I will store valid and invalid email addresses based on output of regex implementation  
valid_emails = []
not_valid_emails = []

for email in emails: # iterate through each email address in the list gathered from the data set 
    if re.search("^(\w+[\.\-]?)+@+\w+\.(com|org|edu|net)$", email, re.IGNORECASE):
        
        # Use the re.search function to search for a match with the regular expression pattern in the email string
        # The regular expression pattern is surrounded by quotes and passed as the first argument to the re.search function
        # The email address we're matching against is passed as the second argument, in this case what we're iterating over 
        # The re.IGNORECASE flag is used to make the search case-insensitive

        valid_emails.append(email) #If the search function returns true (aka email passes) then the email gets added to the valid emails list 
        print(f"{email} is a valid email")

    else:
        not_valid_emails.append(email) # If email does not match regex, then that email is added to the not_valid_emails list
        print(f"{email} is NOT a valid email!!")

#What the regualr expression above does:
    #So, the regular expression above is looking for a string that starts with one or more word characters (alphanumeric characters (letters and digits) and Underscores (_)) 
    # or a word character followed by an optional period or hyphen, followed by an @ symbol,followed by one or more word characters, followed by a period, and then either com, org, or edu. 
    # The whole string must match the pattern, and no additional characters can be present at the end.

#The \n denotes that I want a space before and after this print statement is executed to make reading it in the command line easier 
print("\nThe following email addresses are valid and an email will be sent to them saying their specific 3D print is completed and ready for pickup:") 
print(valid_emails)

print("\nThe following email addresses were flagged for being invalid and no email will be sent to them:")
print(not_valid_emails)


#Taking the list of invalid emails and adding that information into the invalid_email column for our updated data frame
# df_filter.loc is a locator method to find the row(s) where email column equals the current invalid email
for invalid_email in not_valid_emails:
    df_filter.loc[df_filter["email"] == invalid_email, "invalid_email"] = "X" #Essentially, when value in email variable is in the not_valid_emails list, we add an "X" in the invalid_email variable to say 
                                                                                #that observation contains an invalid email. 

print("\nThe invalid_email varaible has been updated with the above invalid emails\n")


#filter out invalid emails from the data frame to send out emails that properly correspond to their print  
df_filter_valid = df_filter[(df_filter["invalid_email"] != "X")]
df_filter_valid #Means the data is filtered (columns and observations) + only the valid emails are listed 


#Prompt in the middle of the code making sure that we want to contine sending the emails out to patrons 
import sys

user_input = input("Seeing the valid and invalid email addresses, would you like to continue with sending emails to the valid email addresses? (y/n) ")

if user_input.lower() == "y":
    print("Continuing with the program.\n")
else:
    print("Exiting the program.")
    sys.exit()


import datetime

today_str = datetime.date.today().strftime("%m/%d/%y") #Gets today's date in the format of "02/26/23"

#Converting the date string we got above into an actual date  
today_date = datetime.datetime.strptime(today_str, "%m/%d/%y").date()

#We use the .fillna() method to replace all empty values in patron_contacted with today's date regardless of whether email is valid or not  
df_filter.loc[df_filter['contacted_date'].isna(), 'contacted_date'] = today_date


#Code to send emails to patrons informing them that their specific 3D print is completed and ready for pickup 
from email.message import EmailMessage # Library that formats code and strings as an email 
import ssl # Adds security to our email 
import smtplib # Library that actually sends the email 
import time  #used to get time.sleep function pause the program as to not send emails too fast and get blocked by gmail

email_sender = "egr3dprinting@gmail.com"
email_pass = "vpbpggszhbhnklkz"
email_receivers = valid_emails.copy()

# Extract the print variables into separate lists
prints = df_filter_valid["print"].tolist()

#Defines subject header for emails 
subject = "3D Print is Ready for Pickup!"

# Adding a layer of security to our email 
context = ssl.create_default_context()

for i in range(len(email_receivers)):
    #What we want our message in the email to say 
    body = "Hello," + "\n\n" + f"Your {prints[i]} print has been completed and is ready for pickup!" + \
    "\n\n" + f" Please pick up your {prints[i]} print within 2 weeks at the lower level desk of the EGR KDL during open hours." + \
    "\n\n" + "From the EGR KDL 3D Printing Team"
    
    # Matching up the parameters above for our email to the format python expects them 
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receivers[i]
    em['Subject'] = subject
    em.set_content(body)
    
    #Opening the connection to the gmail sever, logging in and then sending the email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_pass)
        smtp.sendmail(email_sender, email_receivers[i], em.as_string())
        
        #Adding a nice print step confimring an email has been sent to all of the email addresses 
        print("Email successfully sent to " + email_receivers[i])

        # Adds an "X" in the patron_contacted column for the email addresses we just sent messages to
        df_filter.loc[df_filter["email"] == email_receivers[i], "patron_contacted"] = "X"
        
    #Sleeping the program for 2 seconds to avoiding sending too many emails back to back and getting timed out by gmail
    time.sleep(2)

#Print statement to let me know that the program is done     
print("All emails sent successfully!\n")


import os 

#Recreating the folder to save excel files if it doesn't already exist, os.path.expanduser is used because os.makedirs doesn't like the ~
home_dir = os.path.expanduser("~")
#Specifies the directory we're going to search for files and create the new excel file noting patrons contacted 
directory_save = os.path.join(home_dir, "Desktop", "3D Printing output files")
os.makedirs(directory_save, exist_ok=True)

string = "Patron_Contacted_Updated"  #This string will be in the excel file we're searching for to either overwrite it or not  

today = datetime.datetime.now().strftime("%b %d") #Gets today's today in the format of "Feb 18"

file_name = f"Patron_Contacted_Updated {today}.xlsx" #Creating a new excel file with the updated date in the name 

file_path = os.path.join(directory_save, file_name) #Combining the directory and file name to create the file path to more easily write this file location later on 


for filename in os.listdir(directory_save): #loops through all the files in the directory to search for the excel file with "Patron_Contacted_Updated" (if it exists) then update it!
    if string in filename:
        
        file_pathy = os.path.join(directory_save, filename)
        answer = input(f"Old version of this file named {filename} exists, would you like to delete it? (y/n) ")

        if answer == 'y': #Code to remove the old excel file directed 
            os.remove(file_pathy)
            print(f"\'{filename}\' has been deleted\n")

        else:
            print(f"The existing file \'{filename}\' will not be overwritten.")
            sys.exit()
        

#Code to create a new excel doc marking the patrons who have been contacted via email
master_df = pd.read_excel(excel_file_master) #loads in master excel doc from onedrive

# changing the contacted_date variable from a datetime to just a date
master_df['contacted_date'] = master_df['contacted_date'].dt.date

master_df.update(df_filter) #updates master_df doc with the updated patron_contacted varaible from df_filter (AKA tallying patrons who have been contacted)

master_df.to_excel(file_path , index=False) #Saves the updated master_df to onedrive to be used to update the master excel doc (that the form feeds into)

print(f"New file: \'{file_name}\' has successfully been created!\n")

#After the emails are sent we need to update the patron_contacted column so this will create another excel I'll use to update the master doc 
print(f"Now, use \'{file_name}\' to update the varaibles 'patron_contacted', 'contacted_date', 'invalid_email' in the Master 3D Printing doc!!!")

time.sleep(6)