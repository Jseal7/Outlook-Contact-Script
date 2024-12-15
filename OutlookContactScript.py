import os
import tkinter.font
import win32com.client
from openpyxl import Workbook, load_workbook
import tkinter
from tkinter import PhotoImage, messagebox

def getOutlookCOntacts():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts = outlook.GetDefaultFolder(10).Items
    contacts_list = []

    for contact in contacts:
        contact_info = {}

        contact_info['Name'] = contact.FullName

        contact_info['Email1'] = contact.Email1Address
        contact_info['Email2'] = contact.Email2Address
        contact_info['Email3'] = contact.Email3Address

        contact_info['Business'] = contact.BusinessTelephoneNumber
        contact_info['Home'] = contact.HomeTelephoneNumber
        contact_info['Mobile'] = contact.MobileTelephoneNumber

        contact_info['Address'] = contact.MailingAddress

        contact_info['Company'] = contact.CompanyName
        contact_info['Job Title'] = contact.JobTitle
        
        contacts_list.append(contact_info)

    return contacts_list

#Function that creates a blank excel sheet to be populated. Also displays message whether it was succesful or not and where it was created.
def createExceltemplate():
    try:
        fileLocation = os.getcwd()
        fileName = 'outlook_contacts.xlsx'
        makeExcel(template=True)
        tkinter.messagebox.showinfo("Success", f"Template created successfully!\n\nFile Name: {fileName}\nLocation: {fileLocation}")
    except Exception as e:
        tkinter.messagebox.showerror("Error", f"An error occurred: {str(e)}")

def makeExcel(contacts_list = None, filename = 'outlook_contacts.xlsx', template = False):
    try:
        excelBook = Workbook()
        currSheet = excelBook.active
        currSheet.title = "Outlook Contacts"
        currSheet.append(['Name', 'Email1', 'Email2', 'Email3', 'Business', 'Home', 'Mobile', 'Address', 'Company', 'Job Title'])
        
        if contacts_list and not template:
            for contact in contacts_list:
                currSheet.append([contact['Name'], contact['Email1'], contact['Email2'], contact['Email3'], contact['Business'], contact['Home'], contact['Mobile'], contact["Address"], contact['Company'], contact['Job Title']])
        
        excelBook.save(filename)

        if not template:
            fileLocation = os.getcwd()
            fileName = 'outlook_contacts.xlsx'
            tkinter.messagebox.showinfo("Success", f"File {fileName} populated with Outlook Contacts!\n\nLocation: {fileLocation}")

    except Exception as e:
        tkinter.messagebox.showerror("Error", f"An error occurred: {str(e)}")

def makeContacts(filename, contacts_list):
    if not os.path.exists(filename):
        fileName = 'outlook_contacts.xlsx'
        tkinter.messagebox.showinfo("No Excel Sheet", f"No file {fileName} exists to populate contacts. Create a \'Template\' and fill in contact info first!")
        return

    excelBook = load_workbook(filename)
    currSheet = excelBook.active

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    contacts = outlook.GetDefaultFolder(10).Items

    nameArr = []
    for i in range(len(contacts_list)):
        nameArr.append(contacts_list[i]['Name'])
    
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

#Function to create the GUI Window for usability.
def makeGui():
    guiWindow = tkinter.Tk()
    guiWindow.title("Excel to Outlook Contact")
    guiWindow.geometry("250x310")
    guiWindow.configure(background='#C0AFE2')

    frame = tkinter.Frame(guiWindow, padx=20, pady=30)
    frame.pack(padx=5, pady=5)
    frame.configure(background='#fff')

    frame.grid_rowconfigure(0, weight=1)
    frame.grid_rowconfigure(1, weight=1)
    frame.grid_rowconfigure(2, weight=1)
    frame.grid_rowconfigure(3, weight=1)
    frame.grid_rowconfigure(4, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    #Create a base directory for where the script is running
    baseDirectory = os.path.dirname(os.path.abspath(__file__))

    #Create the two image paths corresponding to the images for each scripts buttons.
    excelToContactImagePath = os.path.join(baseDirectory, "images", "contact-book.png")
    excelToContactImagePath = excelToContactImagePath.replace("\\", "/")
    contactToExcelImagePath = os.path.join(baseDirectory, "images", "database.png")
    contactToExcelImagePath = contactToExcelImagePath.replace("\\", "/")

    #Changing and resizing into usable image.                       
    excelToContactImage = (PhotoImage(file=excelToContactImagePath)).subsample(10, 10)
    contactToExcelImage = (PhotoImage(file=contactToExcelImagePath)).subsample(8, 8)

    scriptTitle = tkinter.Label(
        frame,
        text="Contact Functions",
        background="#fff", fg='#000000',
        font=('Tahoma', 15, 'bold'))
    scriptTitle.grid(row=0, column=0, pady=10)

    secondRow = tkinter.Frame(frame, background="#fff")
    secondRow.grid(row=1, column=0)

    excelToOutlookButton = tkinter.Button(
        secondRow,
        image = excelToContactImage,
        width = 60,
        height = 65,
        border = 4,
        background='#C0AFE2',
        activebackground='#D6B4FC',
        highlightthickness=2,
        highlightbackground='#02D7FF',
        highlightcolor='#FFFFFF',
        cursor='hand2',
        command=lambda: makeContacts('outlook_contacts.xlsx', getOutlookCOntacts()))
    excelToOutlookButton.pack(side=tkinter.LEFT, padx=10, pady=7)

    outlookToExcelButtton = tkinter.Button(
        secondRow, 
        image = contactToExcelImage,
        background='#C0AFE2',
        width = 60,
        height = 65,
        border = 4,
        activebackground='#D6B4FC',
        highlightthickness=2,
        highlightbackground='#02D7FF',
        highlightcolor='#FFFFFF',
        cursor='hand2',
        command=lambda: makeExcel(getOutlookCOntacts(), 'outlook_contacts.xlsx'))
    outlookToExcelButtton.pack(side=tkinter.LEFT, padx=10, pady=7)

    thirdRow = tkinter.Frame(frame, background="#fff")
    thirdRow.grid(row=2, column=0)

    excelToOutlookLabel = tkinter.Label(
        thirdRow,
        text="Create\nContacts",
        height=3,
        background="#fff", fg='#000000',
        font=('Tahoma', 8, 'bold'))
    excelToOutlookLabel.pack(side=tkinter.LEFT, padx=17, pady=5)

    outlookToExcelLabel = tkinter.Label(
        thirdRow,
        text="Populate\nSheet",
        height=3,
        background="#fff", fg='#000000',
        font=('Tahoma', 8, 'bold'))
    outlookToExcelLabel.pack(side=tkinter.LEFT, padx=17, pady=5)

    thirdRow = tkinter.Frame(frame, background="#fff")
    thirdRow.grid(row=3, column=0, pady=8)

    templateButton = tkinter.Button(
        thirdRow,
        text="Excel Template",
        width=15,
        height=2,
        border=2,
        background="#A7D9A3",
        activebackground='#92D191',
        highlightthickness=2,
        highlightbackground='#02D7FF',
        highlightcolor='#FFFFFF',
        cursor='hand2',
        font=('Tahoma', 8, 'bold'),
        command=lambda: createExceltemplate())
    templateButton.pack(side=tkinter.LEFT, padx=10)

    bottomRow = tkinter.Frame(frame, background="#fff")
    bottomRow.grid(row=4, column=0, pady=8)

    quitButton = tkinter.Button(
        bottomRow,
        text="Quit",
        width = 9,
        height = 2,
        border = 2,
        background="#FF7F7F",
        activebackground='#D50101',
        highlightthickness=2,
        highlightbackground='#02D7FF',
        highlightcolor='#FFFFFF',
        cursor='hand2',
        font=('Tahoma', 8, 'bold'),
        command=guiWindow.quit)
    quitButton.pack(side=tkinter.LEFT, padx=10)

    guiWindow.mainloop()


if __name__ == "__main__":
    desktopPath = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    os.chdir(desktopPath)
    makeGui()