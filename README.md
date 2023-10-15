# Workflow for Recieving, Handling, and Contacting Patron's 3D Print Requests 
 
## Overview 

The 3D printing program used to rely soley on pen and paper for taking in and processing 3D print requests. Once a print request was completed, staff would have to manually contact the patron by phone or email (whichever one was more legible). This was an arduous process that ate up a ton of staff hours that could have been devoted to tasks actually in the job description of a librarian. These paper slips would commonly get lost leading to awkward interactions when a patron came to check on their print request but staff could not find any evidence of it. A typical problem when using pen and paper to recieve 3D print requests was that staff had no way to control the types of prints patrons would request. This lead to many patron requests that simply did not fit the parameters of our 3D printers or were unable to be found on open source repositiroies. In addition, staff had no way to compile useful information on how well the program was performing like most popular requests or colors or simply the turnaround time from recieving a 3D print request to contacting the patron that it was finished.

To fix those issues, I worked to devise a new 3D printing workflow that would leverage an online form and automation to streamline the 3D printing process for both staff and patrons. For the first step, I created an online form that would ask basic contact information and the patrons library card number (in order to better enforce our 1 3D print per month policy). The form gave patrons the option of colored filament to use for the 20 handpicked 3D print rquests tested by staff that we knew worked well with our printers and could be scaled by printing multiple at a time. If a patron wanted something not included in our curated list a link was provided to an online 3D printing respository. Responses from the form were then automatically fed into a shared excel spreadsheet were I built 3 python programs to organize all the data to be easily acted upon by staff. 

No knowledge of coding or the command prompt is requered to utilize each of these programs! Once downloaded as an executable to your PC each program walks/infomrs staff through the process with simple y/n (yes/no) prompts. The programs show staff previews of excel sheets and email lists to verify and proceed if desired. These programs are also designed to ask staff if they wish to delete old versions of excel sheets made by these programs to ensure no valuable data is lost.     

## Emailing Sheet 

To make the printing process more efficient I organzied the 3D prints to be completed list pimrarily by color then by date. This way we could start multiple same colored prints on the printer at once to save time from constatnly switching out the filament color. The program then emails that excel sheet out to 3D print staff with instructions to appropriatly format it to be printed off.    

## Emailing Patrons 

## Expired Prints 
