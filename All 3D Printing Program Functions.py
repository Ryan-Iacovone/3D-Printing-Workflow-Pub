import pandas as pd 
import os as os
import datetime as dt
import re
import time # used to get time.sleep function pause the program as to not send emails too fast and get blocked by gmail
import sys
import requests
import json
from email.message import EmailMessage # formats code and strings as an email 
import ssl # adds security to our email 
import smtplib # actually sends the email 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart 

#Checks which python version is running
#print(sys.executable)

# Sending email to patrons that their print is ready for pickup

def send_completed_prints_email(df):    

# filtering data by the people who's print is completed, they have not been contacted yet, they do not have an invalid email address, and there print has not been picked up yet

    df_filter = df[(df["print_completed"] == "X") & (df["patron_contacted"].isna()) & (df["invalid_email"].isna()) & (df["picked_up"].isna())].copy() #can do is.na or set patron_contacted = "X"

    # Following code applies regex to a list of email addresses we gathered above then sorts by validness 

    # Gathering the list of emails from the filtered data and stripping all whitespaces
    df_filter["email"] = df_filter["email"].str.strip()
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

    print("\nThe invalid_email variable has been updated with the above invalid emails\n")


    #filter out invalid emails from the data frame to send out emails that properly correspond to their print  
    df_filter_valid = df_filter[(df_filter["invalid_email"] != "X")]
    df_filter_valid #Means the data is filtered (columns and observations) + only the valid emails are listed 

    #Prompt in the middle of the code making sure that we want to contine sending the emails out to patrons 

    user_input = input("Seeing the valid and invalid email addresses, would you like to continue with sending emails to the valid email addresses? (y/n) ")

    if user_input.lower() == "y":
        print("Continuing with the program.\n")
    else:
        print("Exiting the program.")
        sys.exit()

    #Gets today's date in the format of "02/26/23"
    today_str = dt.date.today().strftime("%m/%d/%y") 

    #Converting the date string we got above into an actual date  
    today_date = dt.datetime.strptime(today_str, "%m/%d/%y").date()

    #We use the .fillna() method to replace all empty values in patron_contacted with today's date regardless of whether email is valid or not  
    df_filter.loc[df_filter['contacted_date'].isna(), 'contacted_date'] = today_date

    email_sender = ""
    email_pass = ""
    email_receivers = valid_emails.copy()

    # Extract the print variables into separate lists
    prints = df_filter_valid["print"].tolist()

    # Defines subject header for emails 
    subject = "Your 3D Print is Ready for Pickup! 🎨"

    # Adding a layer of security to our email 
    context = ssl.create_default_context()

    def create_html_content(print_name):
        """Creates HTML content for the email"""
        return f"""
        <html>
            <body style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <!-- Replace this URL with your actual logo URL -->
                    <img src="https://ci3.googleusercontent.com/meips/ADKq_NY0zXhTLuBuwPCnp79gq2EjXvc-xSlXkzaTuJnNHzccHKgh2AEmk0CJaK38PcAKHSWI1O5r4-0eyYTFU4R3BebtNE-jVMeW5YLBT56SxNyDnz38ZF35LXMF8WMjC1jt69ziAHgvXhKNtkBqpZBMe01Oav5jMGKl1tYjeb2ubPPEyKUczAPErBPmEqa43mdErCpBWQ28NvL-tlhYIcmU4rzdq4NBrmVbwd3q9RrOSVnkeJY=s0-d-e1-ft#https://messages.librarycustomer.org/images/3848f376-3417-4834-b629-9714b7d7f3a2/MS2X1L2592024/KDL%25252525252525252520Email%25252525252525252520Header.jpg" alt="Library Logo" style="max-width: 480px;">
                </div>
                
                <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; margin-bottom: 20px;">
                    <h1 style="color: #1e196b; text-align: center;">Your 3D Print is Ready! 🎨</h1>
                    
                    <p style="color: #34495e; font-size: 16px;">Hello!</p>
                    
                    <p style="color: #34495e; font-size: 16px;">
                        We've got great news! Your <span style="color: #e74c3c; font-weight: bold;">{print_name}</span> print has been 
                        successfully printed and is ready for pickup.
                    </p>
                    
                    <div style="background-color: #ffffff; border-left: 4px solid #3498db; padding: 15px; margin: 20px 0;">
                        <h3 style="color: #2c3e50; margin-top: 0;">Details:</h3>
                        <ul style="color: #34495e; list-style-type: none; padding-left: 0;">
                            <li style="margin-bottom: 10px;">📦 <strong>Item:</strong> {print_name}</li>
                            <li style="margin-bottom: 10px;">📍 <strong>Pickup Location:</strong> Lower Level (Youth) Desk, EGR KDL</li>
                            <li style="margin-bottom: 10px;">⏰ <strong>Pickup Window:</strong> Within 2 weeks from today</li>
                        </ul>
                    </div>
                    
                    <p style="color: #34495e; font-size: 16px;">
                        Please note that your 3D print should be picked up during our open library hours. 
                    </p>
                    
                    <p style="color: #34495e; font-size: 16px;">We hope you'll be delighted with your 3D print creation!</p>
                    
                    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #bdc3c7;">
                        <p style="color: #34495e; font-size: 16px; margin-bottom: 5px;">Best regards,</p>
                        <p style="color: #1e196b; font-weight: bold; margin-top: 0;">The EGR KDL 3D Printing Team</p>
                    </div>
                    
                    <div style="background-color: #e8f4f8; padding: 15px; border-radius: 5px; margin-top: 20px;">
                        <p style="color: #2980b9; margin: 0;">
                            <strong>P.S.</strong> Interested in what else KDL offers? 
                            <a href="https://kdl.org/kdl-beyond-books-collection/" style="color: #3498db;">Check out</a> 
                            our Beyond Books collection!
                        </p>
                    </div>
                </div>
            </body>
            </html>
        """

    def create_plain_text_content(print_name):
        """Creates plain text content for email clients that don't support HTML"""
        return f"""
    Hello,

    Great news! Your {print_name} has been successfully printed and is ready for pickup.

    Details:
    • Item: {print_name}
    • Pickup Location: Lower Level (Youth) Desk, EGR KDL
    • Pickup Window: Within 2 weeks from today

    Please note that your 3D print should be picked up during our open library hours. 

    We hope you'll be delighted with your 3D creation!

    Best regards,
    The EGR KDL 3D Printing Team
        """

    for i in range(len(email_receivers)):
        # Create the email message
        msg = MIMEMultipart('alternative')
        msg['From'] = email_sender
        msg['To'] = email_receivers[i]
        msg['Subject'] = subject
        
        # Create both plain text and HTML versions
        text_part = MIMEText(create_plain_text_content(prints[i]), 'plain')
        html_part = MIMEText(create_html_content(prints[i]), 'html')
        
        # Add both parts to the message
        # The email client will try to render the last part first
        msg.attach(text_part)
        msg.attach(html_part)
        
        # Opening the connection to the gmail server, logging in and then sending the email
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            try:
                smtp.login(email_sender, email_pass)
            except SMTPAuthenticationError as e:
                print("This is probably the worst error you could get ( • ᴖ • ｡) \n")
                print("Email password failed to connect with gmail account for 'egr3dprinting@gmail.com'. \n")
                print("Please refer to video X on how to fix this issue or contact Ryan. \n")
                print("No emails were sent to patrons. \n")
                print("Program now shutting down... \n")
                time.sleep(15)
                sys.exit()
                
            smtp.sendmail(email_sender, email_receivers[i], msg.as_string())
            
            # Adding a nice print step confirming an email has been sent
            print("Email successfully sent to " + email_receivers[i])

            # Adds an "X" in the patron_contacted column for the email addresses we just sent messages to
            df_filter.loc[df_filter["email"] == email_receivers[i], "patron_contacted"] = "X"
            
        # Sleeping the program for 2 seconds to avoid sending too many emails back to back
        time.sleep(2)

    print("All emails sent successfully!\n")


    #Recreating the folder to save excel files if it doesn't already exist, os.path.expanduser is used because os.makedirs doesn't like the ~
    home_dir = os.path.expanduser("~")
    #Specifies the directory we're going to search for files and create the new excel file noting patrons contacted 
    directory_save = os.path.join(home_dir, "Desktop", "3D Printing output files")
    os.makedirs(directory_save, exist_ok=True)

    string = "Patron_Contacted_Updated"  #This string will be in the excel file we're searching for to either overwrite it or not  

    today = dt.datetime.now().strftime("%b %d") #Gets today's today in the format of "Feb 18"

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

    # reloads in master excel but renames it for clarity
    master_df = df 

    # changing the contacted_date variable from a datetime to just a date only if there are dates already there
    if master_df['contacted_date'].notna().any():
        # Only convert the non-null values to date
        master_df['contacted_date'] = master_df['contacted_date'].dt.date

    master_df.update(df_filter) #updates master_df doc with the updated patron_contacted varaible from df_filter (AKA tallying patrons who have been contacted)

    master_df.to_excel(file_path , index=False) #Saves the updated master_df to onedrive to be used to update the master excel doc (that the form feeds into)

    print(f"New file: \'{file_name}\' has successfully been created!\n")

    #After the emails are sent we need to update the patron_contacted column so this will create another excel I'll use to update the master doc 
    print(f"Now, use \'{file_name}\' to update the varaibles 'patron_contacted', 'contacted_date', 'invalid_email' in the Master 3D Printing doc!!!")

    time.sleep(10)



# Code to send an updated version of the excel sheet to 3D print team every week
def print_sheet_staff(df):
    #Filtering the data frame by patrons who's print has not ealready been completed and who's print has has not already gotten on the physical sheet  
    df = df[(df["print_completed"] != "X") & (df["sheet_print"] != "X")]

        #Select by dropping 
    #Drop specific variables from our data frame

    df_selected = df.drop(columns= ["patron_contacted", "contacted_date", "invalid_email", "print_expired", "sheet_print", "recontacted?", "notes"])

    #Sorting the base data frame primarily by color and then by date for easy read ability 

    df = pd.DataFrame(df)
    df_selected_sorted = df_selected.sort_values(by=['color', 'completion_time'],  ascending= [False, True])


    #simplifying the datetime variable completion_time to make it more readable to staff 

    df_selected_sorted['completion_time'] = pd.to_datetime(df_selected_sorted['completion_time'])
    df_selected_sorted['completion_time'] = df_selected_sorted['completion_time'].dt.strftime("%b %d, %Y") #puts date in the format: "Feb 18, 2023"


    #Initiating a prompt in the program to see if the user wants to continue with creating an excel file 

    print("\nBelow are the first 5 observations of the '3D Prints to be Completed' DataFrame:\n")
    print(df_selected_sorted.head(5))

    while True:
        answer = input("\nSeeing the above DataFrame, would you like to make it an excel file and save it locally? (y/n) ")
        if answer == "y":
            break
        elif answer == "n":
            print("The program will now cease running.")
            time.sleep(5)
            sys.exit()
        else:
            print("Invalid input. Please enter 'y' to continue or 'n' to exit.")


    #Changing the card_number column to a string so it'll appear without sccientific notation in the excel sheet
    df_selected_sorted['card_number'] = df_selected_sorted['card_number'].astype(str)


    #The following code saves the df_selected_sorted data frame I've edited and saves it locally as an excel doc 
    today = dt.datetime.now().strftime("%b %d") #gets the current datatime at the moment the code is run, formats it with strftime as 'Feb 20'

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
    today = dt.datetime.now().strftime("%b %d") #Gets today's today in the format of "Feb 18"
    file_name2 = f"Updated sheet_print {today}.xlsx" #Creating a new excel file with the updated date in the name 
    file_path2 = os.path.join(directory, file_name2) #Combining the directory and file name to create the file path to more easily write this file location later on 

    #we need to read in the orginal file again to a dataframe to get the structure of it making copy and pasting the inputs back in way easier 
    #Code to create a new excel doc marking the patrons who have been contacted via email
    master_df = df 

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

    # Another prompt now that the excel file is saved locally seeing if we want to email it out to 3D print staff  
    user_input = input(f"Now that {file_name} is saved locally, would you like to email it out to 3D print staff? (y/n) ")

    if user_input.lower() == "y":
        print("Prepping email to print to staff!\n")
    else:
        print("With the file just saved locally, make sure to print it out!")
        time.sleep(4)
        sys.exit()

    #Code to generate a random quote from the API link to add at the end of the weekly email

    def get_random_quote():
        url = "http://api.quotable.io/random"
        response = requests.get(url)
        if response.status_code == 200:
            quote = response.json()
            return quote["content"] + " - " + quote["author"]
        else:
            return "Error: Could not retrieve quote."

    quote_of_the_week = 'Quote of the week: "' + get_random_quote() + '"'

    #Code to send an email to multiple email addresses at once with an excel attachment   

    #List of emails for 3D print team 
    team_emails = []

    email_sender = ""
    email_pass = ""

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
    em['To'] = team_emails
    em['Subject'] = subject
    em.set_content(body)
    em.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename= file_name)

    # Sending an email to all KDL 3D print staff all at once
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_pass)
        smtp.sendmail(email_sender, team_emails, em.as_string())

    print("\nEmail successfully sent to printing team!")
    time.sleep(3)


# Recontacting patrons
def recontacting_patrons(df):

    # calculate the date two weeks ago
    two_weeks_ago = dt.datetime.now() - dt.timedelta(days=14)

    # filtering data by people who have been contacted more than two weeks ago but not yet pick up their print
    df_filter = df[(df["print_completed"] == "X") & (df["patron_contacted"] == "X") & (df["invalid_email"].isna()) & (df["picked_up"].isna()) & (df["recontacted?"].isna()) & (df["contacted_date"] <= two_weeks_ago)].copy() 

    # Stripping all whitespaces in the email variable 
    df_filter["email"] = df_filter["email"].str.strip()

    # Displaying emails and names we will recontact to user:
    print("Here's the list of patrons who haven't picked up their print in 2 weeks and will therefore be recontacted: \n")
    display_df = df_filter[["ID" ,"contacted_date", "name", "email",  "print"]]
    print(display_df, "\n")

    while True:
        answer = input("\nSeeing the above DataFrame, would you like to continue sending out recontacting emails? (y/n) ")
        if answer == "y":
            break
        elif answer == "n":
            print("The program will now cease running.")
            time.sleep(5)
            sys.exit()
        else:
            print("Invalid input. Please enter 'y' to continue or 'n' to exit.")

    #Code to send emails to patrons informing them that their specific 3D print is completed and ready for pickup 
    email_sender = ""
    email_pass = ""
    email_receivers = df_filter["email"].tolist() # Gathering the list of emails from the filtered data

    # Extract the print variables into separate lists
    prints = df_filter["print"].tolist()

    #Defines subject header for emails 
    subject = "REMINDER: 3D Print is Still Ready for Pickup!"

    # Adding a layer of security to our email 
    context = ssl.create_default_context()

    for i in range(len(email_receivers)):
        #What we want our message in the email to say 
        body = "Hello," + "\n\n" + f"Your {prints[i]} print has been completed and is still ready for pickup!" + \
        "\n\n" + f" Please pick up your {prints[i]} print as soon as possible at the lower level desk of the EGR KDL during open hours." + \
        "\n\n" + "From the EGR KDL 3D Printing Team"
        
        # Matching up the parameters above for our email to the format python expects them 
        em = EmailMessage()
        em['From'] = email_sender
        em['To'] = email_receivers[i]
        em['Subject'] = subject
        em.set_content(body)
        
        #Opening the connection to the gmail sever, logging in and then sending the email
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            try:
                smtp.login(email_sender, email_pass)
            except SMTPAuthenticationError as e:
                print("This is probably the worst error you could get ( • ᴖ • ｡) \n")
                print("Email password failed to connect with gmail account for 'egr3dprinting@gmail.com'. \n")
                print("Please refer to video X on how to fix this issue or contact Ryan. \n") # NEED TO FIGURE OUT HOW THEY'RE GOING TO HANDEL THIS ERROR
                print("No emails were sent to patrons. \n")
                print("Program now shutting down... \n")
                time.sleep(15)
                sys.exit()
                
            smtp.sendmail(email_sender, email_receivers[i], em.as_string())
            
            #Adding a nice print step confimring an email has been sent to all of the email addresses 
            print("Recontacted email successfully sent to " + email_receivers[i])

            # Adds an "X" in the patron_contacted column for the email addresses we just sent messages to
            df_filter.loc[df_filter["email"] == email_receivers[i], "recontacted?"] = "X"
            
        #Sleeping the program for 2 seconds to avoiding sending too many emails back to back and getting timed out by gmail
        time.sleep(2)

    #Print statement to let me know that the program is done     
    print("\nAll recontacted emails successfully sent!\n")


    #Recreating the folder to save excel files if it doesn't already exist, os.path.expanduser is used because os.makedirs doesn't like the ~
    home_dir = os.path.expanduser("~")
    #Specifies the directory we're going to search for files and create the new excel file noting patrons contacted 
    directory_save = os.path.join(home_dir, "Desktop", "3D Printing output files")
    os.makedirs(directory_save, exist_ok=True)

    string = "Patron_Recontacted_Updated"  #This string will be in the excel file we're searching for to either overwrite it or not  

    today = dt.datetime.now().strftime("%b %d") #Gets today's today in the format of "Feb 18"

    file_name = f"Patron_Recontacted_Updated {today}.xlsx" #Creating a new excel file with the updated date in the name 

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

    #reloading in master excel doc 
    master_df = df 

    # changing the contacted_date variable from a datetime to just a date only if there are dates already there
    if master_df['contacted_date'].notna().any():
        # Only convert the non-null values to date
        master_df['contacted_date'] = master_df['contacted_date'].dt.date

    master_df.update(df_filter) #updates master_df doc with the updated patron_contacted varaible from df_filter (AKA tallying patrons who have been contacted)

    master_df.to_excel(file_path , index=False) #Saves the updated master_df to onedrive to be used to update the master excel doc (that the form feeds into)

    print(f"New file: \'{file_name}\' has successfully been created!\n")

    #After the emails are sent we need to update the patron_contacted column so this will create another excel I'll use to update the master doc 
    print(f"Now, use \'{file_name}\' to update the  'recontacted?' variable in the Master 3D Printing doc!!!")

    time.sleep(10)


#Importing in master 3D printing data set from onedrive to get updated form data 
try:
    excel_file_master = r"~\Desktop\EGR KDL Master 3D Printing List.xlsx"
    df = pd.read_excel(excel_file_master)
    print("'EGR KDL Master 3D Printing List.xlsx' file found and now being imported! \n")
except FileNotFoundError:
    print("EGR KDL Master 3D Printing List not found. \n")
    print("Please save 'EGR KDL Master 3D Printing List.xlsx' to your desktop for this script to work.")
    print("Program now shutting down...")
    time.sleep(15)
    sys.exit()

# Checking to ensure all the variables are found the df 
required_variables = ['ID', 'completion_time', 'card_number', 'name', 'email', 'phone_number', 'print', 'color', 'caregiver_name', 'print_started', 'print_completed', 'patron_contacted', 'contacted_date', 'recontacted?', 
    'invalid_email','picked_up', 'print_expired', 'failed', 'sheet_print', 'notes']

def check_df_variables(df, required_variables):
    try:
        # Check if all required variables are present in the DataFrame
        missing_variables = [var for var in required_variables if var not in df.columns]
        if missing_variables:
            raise ValueError(f"Missing or mispelled variables: {','.join(missing_variables)}. \n\nPlease ensure that those variables listed above are in the excel sheet and spelled as shown. \n\nThese are all the variables, their correct spelling, and proper order that should be in this excel sheet: {required_variables}")

        else:
            print(f"'EGR KDL Master 3D Printing List.xlsx' successfully loaded in! Program procceding...\n")
    except Exception as e:
        print(f"ERROR OCCURED: {e}")
        time.sleep(20)
        sys.exit()

check_df_variables(df, required_variables)

# Initiate the program where staff picks which function to execute
while True:
    u_input = input("""Hello, what you like to do today? \n
    1. Send Emails to patrons that their 3D prints are completed.
    2. Create an excel of prints that still need to be printed.
    3. Send a follow up email to patrons who's prints are completed and have already been contacted more than 2 weeks ago. \n
    Please select type either 1, 2, or 3 or type 'n' to exit program. \n""").strip()

    if u_input == "1":
        print("Initializing program to contact patrons that their 3D print is completed and ready for pickup \n")
        time.sleep(2)
        send_completed_prints_email(df)
        break
    elif u_input == "2":
        print("Initializing program to print off excel sheet with prints to be completed \n")
        time.sleep(2)
        print_sheet_staff(df)
        break
    elif u_input == "3":
        print("Initializing program to recontact patrons informing them they still need to pick up their 3D print request \n")
        time.sleep(2)
        recontacting_patrons(df)
        break
    elif u_input == "n":
        print("Exiting program. Shutting down...")
        time.sleep(6)
        sys.exit()
    else:
        print("Invalid input. Please either '1', '2', or '3 to continue or 'n' to exit.")