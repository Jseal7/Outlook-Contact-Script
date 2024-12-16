import os
import tkinter.font
import win32com.client
from openpyxl import Workbook, load_workbook
import tkinter
from tkinter import PhotoImage, messagebox
import customtkinter

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
        currSheet.append(['Name', 'Email1', 'Email2', 'Email3', 'Business Phone', 'Home Phone', 'Mobile Phone', 'Address', 'Company', 'Job Title'])
        
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
    guiWindow = customtkinter.CTk(fg_color='#49332b')
    guiWindow.title("Excel to Outlook Contact")
    guiWindow.geometry("240x355")

    frame = tkinter.Frame(guiWindow, padx=20, pady=30)
    frame.pack(padx=5, pady=5)
    frame.configure(background='#edd8bc')

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

    scriptTitle = customtkinter.CTkLabel(
        frame,
        text="Contact Functions",
        fg_color="#edd8bc",
        text_color='#49332b',
        font=('Arial', 21, 'bold'))
    scriptTitle.grid(row=0, column=0, pady=10)

    secondRow = tkinter.Frame(frame, background="#edd8bc")
    secondRow.grid(row=1, column=0)

    excelToOutlookButton = tkinter.Button(
        secondRow,
        image = excelToContactImage,
        width = 60,
        height = 65,
        border = 4,
        background='#66473c',
        foreground='#000',
        activebackground='#D6B4FC',
        cursor='hand2',
        command=lambda: makeContacts('outlook_contacts.xlsx', getOutlookCOntacts()))
    excelToOutlookButton.pack(side=tkinter.LEFT, padx=10, pady=7)

    outlookToExcelButtton = tkinter.Button(
        secondRow, 
        image = contactToExcelImage,
        background='#66473c',
        width = 60,
        height = 65,
        border = 4,
        activebackground='#D6B4FC',
        cursor='hand2',
        command=lambda: makeExcel(getOutlookCOntacts(), 'outlook_contacts.xlsx'))
    outlookToExcelButtton.pack(side=tkinter.LEFT, padx=10, pady=7)

    thirdRow = customtkinter.CTkFrame(frame, fg_color="#edd8bc")
    thirdRow.grid(row=2, column=0)

    excelToOutlookLabel = customtkinter.CTkLabel(
        thirdRow,
        text="Create\nContacts",
        height=3,
        fg_color='#edd8bc',
        text_color='#49332b',
        font=('Arial', 12, 'bold'))
    excelToOutlookLabel.pack(side=tkinter.LEFT, padx=17, pady=5)

    outlookToExcelLabel = customtkinter.CTkLabel(
        thirdRow,
        text="Populate\nSheet",
        height=3,
        fg_color='#edd8bc',
        text_color='#49332b',
        font=('Arial', 12, 'bold'))
    outlookToExcelLabel.pack(side=tkinter.LEFT, padx=17, pady=5)

    fourthRow = customtkinter.CTkFrame(frame, fg_color="#edd8bc")
    fourthRow.grid(row=3, column=0, pady=(35,0))

    templateButton = customtkinter.CTkButton(
        fourthRow,
        text="Template",
        width=60,
        height=35,
        border_color='#49332b',
        border_width=3,
        fg_color='#edd8bc',
        hover_color='#a87765',
        text_color='#49332b',
        corner_radius=15,
        cursor='hand2',
        font=('Arial', 11, 'bold'),
        command=lambda: createExceltemplate())
    templateButton.pack(side=tkinter.LEFT, padx=7)

    quitButton = customtkinter.CTkButton(
        fourthRow,
        text="Quit",
        width = 60,
        height = 35,
        fg_color='#edd8bc',
        hover_color='#f73135',
        border_color='#850918',
        border_width=3,
        text_color='#850918',
        corner_radius=15,
        cursor='hand2',
        font=('Arial', 11, 'bold'),
        command=guiWindow.quit)
    quitButton.pack(side=tkinter.LEFT, padx=7)

    guiWindow.mainloop()


if __name__ == "__main__":
    desktopPath = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    os.chdir(desktopPath)
    makeGui()