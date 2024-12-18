<h1>Outlook Contacts Manager</h1>
This Python script, with an added Graphic User Interface (GUI), allows contact management between the Outlook contact book and Excel. Its functionalities include exporting Outlook Contacts into a formatted Excel Sheet, populating an Outlook Contact book from a user-filled-out Excel sheet, and creating the Excel template a user fills out. The project was built using Python and the GUI library CustomTkinter.

<h1>Setup Instructions</h1>
Step 1: Clone or Download the Repository<br/>

 * Use: https://github.com/Jseal7/Outlook-Contact-Script.git<br/>

 * Move to file: cd Outlook-Contact-Script<br/>

<br/>
Step 2: Ensure Directory Structure<br/>
<pre>
Outlook-Contact-Script/
|--> images/
|  |--> contact-book.png
|  |--> database.png
|  |--> quit.png
|  |--> template.png
|--> OutlookContactScript.py
|--> README.md
 </pre>

Step 3: Install Dependencieses<br/>
 - pip install pywin32 openpyxl pillow customtkinter<br/>

Step 4: Run Script<br/>
 - python OutlookContactScript.py<br/>

<h1>GUI Functions</h1>

- Create Contacts:  
  This function grabs contact details from an Excel sheet and populates users' Outlook Contact book with contacts.

- Populate Sheet:  
  This function takes contact information from each Outlook Contact and puts it in a formatted Excel sheet, where each row represents an individual contact.

- Create Template:  
  Creates the basic formatted Excel sheet to fill in contact information before creating Outlook contacts.

- Quit Program:  
  Safely exits GUI.

<h1>Troubleshooting</h1>

**Images Not Displayed:** Double-check that images are stored in the 'images' file and follow the correct file structure.<br/>
**Dependencies Missing:** Install the above dependencies using pip in the terminal.<br/>
**Outlook/Excel Error:** Ensure Outlook and Excel are properly installed and set up following Microsoft's login process.<br/>
**Unable to Find Excel Sheet:** When the Excel is created it should default to your desktop. If not, search for a file named "outlook_contacts.xlsx".<br/>


<h1>Excel Sheet Format Example</h1>

Example of the data that would be inputted and taken to create a contact.

| Name          | Email1             | Email2          | Email3          | Business Phone | Home Phone | Mobile Phone | Address        | Company       | Job Title   |
|---------------|--------------------|-----------------|-----------------|----------------|------------|--------------|----------------|---------------|-------------|
| John Doe      | johnD@gmail.com    | john.alt@abc.com|                 | 1234567890     |            | 9876543210   | 123 Elm St.    | ABC Ltd.      | Manager     |
| Jane Smith    | janeS@gmail.com    |                 |                 |                | 5555551234 |              | 456 Oak Ave.   | XYZ Inc.      | Engineer    |

<h1>Screenshot of GUI</h1>

![image](https://github.com/user-attachments/assets/cf5bfc23-bc70-4a13-87a8-64d84c8fb214)



