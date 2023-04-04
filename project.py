import openpyxl

fileName = "Book1.xlsx"

workBook = openpyxl.load_workbook(fileName)
sheet = workBook.active
columnList= ["A", "B", "C"]
rowCount = sheet.max_row


def emailValidation (firstName, lastName, email):
    firstName = str(firstName).lower().replace(" ", "")
    lastName = str(lastName).lower().replace(" ", "")

    basicPattern = firstName + "." + lastName + "@georgiancollege.ca"

    # creates email if the field is empty
    if(email == None):
        print("Email cell", str(column) + str(row), "was empty so a new email was generated automatically based on the cell information from the same row")
        email = basicPattern
    
    # cleans the email from spaces
    email = email.replace(" ","").lower()

    # checks if email is the same as required pattern
    if (email != basicPattern):
        print("There is something wrong with the email in the cell", str(column) + str(row))

    return email

for row in range(2, rowCount + 1):
    for column in columnList:
        cellContent = sheet[str(column)+str(row)]

        if (cellContent == None and column != "C"):
            print("Cell", str(column) + str(row), "is empty")
            # sheet.delete_row(row, 1)
            # workBook.save(fileName)
        
        if (column == "A"):
            firstName = cellContent.value
        elif (column == "B"): 
            lastName = cellContent.value
        elif (column == "C"):
            email = cellContent.value
            
    email = emailValidation(firstName, lastName, email)
    sheet["C" + str(row)] = email
    workBook.save(fileName)
    
workBook.close()