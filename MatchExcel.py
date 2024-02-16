import glob
import Mailing as mailing
import openpyxl

hl = openpyxl.load_workbook("abc")  # This opens the excel file
abu = hl.active  # This activates the excel file

path = "abc"  # This locates the pdf files

for file in glob.glob(path + "/*.pdf"):  # This looks at the folder and find all the files that has pdf extension.
    if file.endswith(".pdf"):
        print(file)
        for x in range(1, 400):
            Last_Name_loc = "B" + str(x)  # This writes the cell location in the database
            First_Name_loc = "A" + str(x)
            Banner_ID_loc = "C" + str(x)
            E_mail_loc = "D" + str(x)
            firstName = abu[First_Name_loc].value  # This gets the first name value from the location
            fullPath = path + "/" + str(
                abu[Banner_ID_loc].value) + ".pdf"  # This changes the banner ID in spreadsheet to .pdf name
            if fullPath == file:
                print(firstName)
                print(abu[E_mail_loc].value)
                mailing.SendEmail(file, firstName, abu[E_mail_loc].value)
