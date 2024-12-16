import os
import tkinter.font
import win32com.client
from openpyxl import Workbook, load_workbook
import tkinter
import customtkinter
from PIL import Image

def getOutlookContacts():
    """
    Retrieves contacts from users Outlook Contact Book.

    Returns:
        list: List of Dictionaries containing contacts information.  
    """

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts = outlook.GetDefaultFolder(10).Items
    contactsList = []

    for contact in contacts:
        contactInfo = {}

        contactInfo['Name'] = contact.FullName

        contactInfo['Email1'] = contact.Email1Address
        contactInfo['Email2'] = contact.Email2Address
        contactInfo['Email3'] = contact.Email3Address

        contactInfo['Business'] = contact.BusinessTelephoneNumber
        contactInfo['Home'] = contact.HomeTelephoneNumber
        contactInfo['Mobile'] = contact.MobileTelephoneNumber

        contactInfo['Address'] = contact.MailingAddress

        contactInfo['Company'] = contact.CompanyName
        contactInfo['Job Title'] = contact.JobTitle
        
        contactsList.append(contactInfo)

    return contactsList

def createExceltemplate():
    """
    Creates a blank excel sheet to be populated.

    Returns:
        messagebox: Displays the template creation was successful with the files name and location.

    Raises:
        Exception: If any error occurs that prevents the template creation.
    """

    try:
        fileLocation = os.getcwd()
        fileName = 'outlook_contacts.xlsx'
        makeExcel(template=True)
        tkinter.messagebox.showinfo("Success", f"Template created successfully!\n\nFile Name: {fileName}\nLocation: {fileLocation}")
    except Exception as e:
        tkinter.messagebox.showerror("Error", f"An error occurred: {str(e)}")


def makeExcel(contactsList = None, fileName = 'outlook_contacts.xlsx', template = False):
    """
    Makes an excel sheet and populates it with contacts if it is not marked as a template.

    Args:
        contactsList: A list of dictionaries containing contacts information. Default set to None.
        fileName: The file name for excel sheet. Default set to 'outlook_contacts.xlsx'.
        template: Boolean for if the excel sheet with be a template or not. Default set to False.

    Returns:
        messagebox: Creates a formatted excel sheet and returns its success status, the file name, and where the file was stored.

    Raises:
        Exception: If any error occurs that prevents making a new excel sheet.
    """

    try:
        excelBook = Workbook()
        currSheet = excelBook.active
        currSheet.title = "Outlook Contacts"
        currSheet.append(['Name', 'Email1', 'Email2', 'Email3', 'Business Phone', 'Home Phone', 'Mobile Phone', 'Address', 'Company', 'Job Title'])
        
        if contactsList and not template:
            for contact in contactsList:
                currSheet.append([contact['Name'], contact['Email1'], contact['Email2'], contact['Email3'], contact['Business'],
                                contact['Home'], contact['Mobile'], contact["Address"], contact['Company'], contact['Job Title']])
        
        excelBook.save(fileName)

        if not template:
            fileLocation = os.getcwd()
            fileName = 'outlook_contacts.xlsx'
            tkinter.messagebox.showinfo("Success", f"File {fileName} populated with Outlook Contacts!\n\nLocation: {fileLocation}")

    except Exception as e:
        tkinter.messagebox.showerror("Error", f"An error occurred: {str(e)}")

def makeContacts(fileName, contactsList):
    """
    Creates contacts in Outlook Contact Book from list of contacts information.

    Args:
        contactsList: A list of dictionaries containing contacts information.
        fileName: The file name for excel sheet.

    Returns:
        messagebox: Creates Outlook contacts and returns its success status and the file name. If a file of the same name already exists,
                    it returns a message to start with a template.

    Raises:
        Exception: If any error occurs that prevents making a new excel sheet.
    """

    if not os.path.exists(fileName):
        fileName = 'outlook_contacts.xlsx'
        tkinter.messagebox.showinfo("No Excel Sheet", f"No file {fileName} exists to populate contacts. Create a \'Template\' and fill in contact info first!")
        return

    excelBook = load_workbook(fileName)
    currSheet = excelBook.active

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts = outlook.GetDefaultFolder(10).Items

    nameArr = []
    for i in range(len(contactsList)):
        nameArr.append(contactsList[i]['Name'])
    
    for i in range(2, currSheet.max_row + 1):
        if (not (currSheet.cell(row= i, column= 1).value in nameArr)):
            contactItem = contacts.Add("IPM.Contact")

            if (currSheet.cell(row= i, column= 1).value != None):
                contactItem.FullName = currSheet.cell(row= i, column= 1).value
            if (currSheet.cell(row= i, column= 2).value != None):
                contactItem.Email1Address = currSheet.cell(row= i, column= 2).value
            if (currSheet.cell(row= i, column= 3).value != None):
                contactItem.Email2Address = currSheet.cell(row= i, column= 3).value
            if (currSheet.cell(row= i, column= 4).value != None):
                contactItem.Email3Address = currSheet.cell(row= i, column= 4).value    
            if (currSheet.cell(row= i, column= 5).value != None):
                contactItem.BusinessTelephoneNumber = currSheet.cell(row= i, column= 5).value
            if (currSheet.cell(row= i, column= 6).value != None):
                contactItem.HomeTelephoneNumber = currSheet.cell(row= i, column= 6).value
            if (currSheet.cell(row= i, column= 7).value != None):
                contactItem.MobileTelephoneNumber = currSheet.cell(row= i, column= 7).value
            if (currSheet.cell(row= i, column= 8).value != None):
                contactItem.MailingAddress = currSheet.cell(row= i, column= 8).value
            if (currSheet.cell(row= i, column= 9).value != None):
                contactItem.CompanyName = currSheet.cell(row= i, column= 9).value
            if (currSheet.cell(row= i, column= 10).value != None):
                contactItem.JobTitle = currSheet.cell(row= i, column= 10).value

            contactItem.Save()

    fileName = 'outlook_contacts.xlsx'
    tkinter.messagebox.showinfo("Success", f"Outlook Contacts were populated from file {fileName}!")

def makeGui():
    """Makes a Graphical User Interface for using the scripts functions."""

    # Establishing of the intial outer GUI window. Instatiates windows size and makes in unresizable.
    guiWindow = customtkinter.CTk(fg_color='#49332b')
    guiWindow.title("Excel to Outlook Contact")
    guiWindow.geometry("250x315")
    guiWindow.resizable(False, False)
    guiWindow.attributes("-fullscreen", False)

    # Creating and configuring the internal box that holds the GUI's content.
    frame = customtkinter.CTkFrame(guiWindow, fg_color='#edd8bc')
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    frame.grid_rowconfigure((0, 1, 2, 3, 4, 5), weight=1)
    frame.grid_columnconfigure(0, weight=1)

    # Create a base directory for where the script is running
    baseDirectory = os.path.dirname(os.path.abspath(__file__))

    # Create the two image paths corresponding to the images for each scripts buttons.
    excelToContactImagePath = os.path.join(baseDirectory, "images", "contact-book.png")
    excelToContactImagePath = excelToContactImagePath.replace("\\", "/")

    contactToExcelImagePath = os.path.join(baseDirectory, "images", "database.png")
    contactToExcelImagePath = contactToExcelImagePath.replace("\\", "/")

    templateImagePath = os.path.join(baseDirectory, "images", "template.png")
    templateImagePath = templateImagePath.replace("\\", "/")

    quitImagePath = os.path.join(baseDirectory, "images", "quit.png")
    quitImagePath = quitImagePath.replace("\\", "/")
    
    # Changing and resizing into usable image.                       
    excelToContactImage = customtkinter.CTkImage(light_image=Image.open(excelToContactImagePath), size=(48, 48))
    contactToExcelImage = customtkinter.CTkImage(light_image=Image.open(contactToExcelImagePath), size=(48, 48))
    templateImage = customtkinter.CTkImage(light_image=Image.open(templateImagePath), size=(48, 48))
    quitImage = customtkinter.CTkImage(light_image=Image.open(quitImagePath), size=(48, 48))

    # Title and Underline (first row start)
    scriptTitle = customtkinter.CTkLabel(
        frame,
        text="Contact Functions",
        fg_color="#edd8bc",
        text_color='#30221d',
        font=('Arial', 22, 'bold'))
    scriptTitle.grid(row=0, column=0, pady=(10, 0))

    titleUnderline = customtkinter.CTkFrame(
        frame,
        fg_color="#30221d",
        height=2,
        width=200)
    titleUnderline.grid(row=1, column=0, sticky="n")
    # (first row end)

    # Top two function buttons (second row start)
    secondRow = customtkinter.CTkFrame(frame, fg_color="#edd8bc")
    secondRow.grid(row=2, column=0)

    excelToOutlookButton = customtkinter.CTkButton(
        secondRow,
        image = excelToContactImage,
        width = 65,
        height = 65,
        border_color = '#49332b',
        border_width = 3,
        fg_color='#c9b69d',
        hover_color='#b5a48d',
        text='',
        cursor='hand2',
        command=lambda: makeContacts('outlook_contacts.xlsx', getOutlookContacts()))
    excelToOutlookButton.pack(side=tkinter.LEFT, padx=10, pady=(7,0))

    outlookToExcelButton = customtkinter.CTkButton(
        secondRow, 
        image = contactToExcelImage,
        width = 65,
        height = 65,
        border_color = '#49332b',
        border_width = 3,
        fg_color='#c9b69d',
        hover_color='#b5a48d',
        text = '',
        cursor = 'hand2',
        command = lambda: makeExcel(getOutlookContacts(), 'outlook_contacts.xlsx'))
    outlookToExcelButton.pack(side=tkinter.LEFT, padx=10, pady=(7,0))
    # (second row end)

    # First two labels depicting button functions (third row start)
    thirdRow = customtkinter.CTkFrame(frame, fg_color="#edd8bc")
    thirdRow.grid(row=3, column=0)

    excelToOutlookLabel = customtkinter.CTkLabel(
        thirdRow,
        text="Create\nContacts",
        height=3,
        fg_color='#edd8bc',
        text_color='#33231e',
        font=('Arial', 12, 'bold'))
    excelToOutlookLabel.pack(side=tkinter.LEFT, padx=17, pady=(0,10))

    outlookToExcelLabel = customtkinter.CTkLabel(
        thirdRow,
        text="Populate\nSheet",
        height=3,
        fg_color='#edd8bc',
        text_color='#33231e',
        font=('Arial', 12, 'bold'))
    outlookToExcelLabel.pack(side=tkinter.LEFT, padx=17, pady=(0,10))
    # (third row end)

    # Second row of buttons (fourth row)
    fourthRow = customtkinter.CTkFrame(frame, fg_color="#edd8bc")
    fourthRow.grid(row=4, column=0)

    templateButton = customtkinter.CTkButton(
        fourthRow,
        image = templateImage,
        width = 65,
        height = 65,
        border_color = '#49332b',
        border_width = 3,
        fg_color='#c9b69d',
        hover_color='#b5a48d',
        text='',
        cursor='hand2',
        command=lambda: createExceltemplate())
    templateButton.pack(side=tkinter.LEFT, padx=10, pady=(7,0))

    quitButton = customtkinter.CTkButton(
        fourthRow,
        image = quitImage,
        width = 65,
        height = 65,
        border_color = '#850918',
        border_width = 3,
        fg_color='#db8f91',
        hover_color='#e36f72',
        text='',
        cursor='hand2',
        command=guiWindow.quit)
    quitButton.pack(side=tkinter.LEFT, padx=10, pady=(7,0))
    # (fourth row end)

    # Second row of labels depicting the functions of the buttons above them. (fifth row start)
    fifthRow = customtkinter.CTkFrame(frame, fg_color="#edd8bc")
    fifthRow.grid(row=5, column=0)

    templateLabel = customtkinter.CTkLabel(
        fifthRow,
        text="Create\nTemplate",
        height=3,
        fg_color='#edd8bc',
        text_color='#33231e',
        font=('Arial', 12, 'bold'))
    templateLabel.pack(side=tkinter.LEFT, padx=17, pady=(0,10))

    quitLabel = customtkinter.CTkLabel(
        fifthRow,
        text="Quit\nProgram",
        height=3,
        fg_color='#edd8bc',
        text_color='#33231e',
        font=('Arial', 12, 'bold'))
    quitLabel.pack(side=tkinter.LEFT, padx=17, pady=(0,10))
    # (fifth row end)

    guiWindow.mainloop()

if __name__ == "__main__":
    desktopPath = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    os.chdir(desktopPath)
    makeGui()