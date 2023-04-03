import openpyxl

fileName = "Book1.xlsx"


workBook = openpyxl.load_workbook(fileName)
sheet = workBook.active
columnList= ["A", "B", "C"]
rowCount = sheet.max_row


def emailValidation (firstName, lastName, email):
    firstName = str(firstName).lower().replace(" ", "")
    lastName = str(lastName).lower().replace(" ", "")

    # creates email if the field is empty
    if(email == None):
        email = firstName + "." + lastName + "@georgiancollege.ca"
    
    
    email = email.replace(" ","").lower()
    return email


for row in range(2, rowCount + 1):
    for column in columnList:
        cellContent = sheet[str(column)+str(row)]

        if (cellContent == None and column != "C"):
            sheet.delete_row(row, 1)
            workBook.save(fileName)

        if (column == "A"):
            firstName = cellContent.value

        if (column == "B"): 
            lastName = cellContent.value
        
        if (column == "C"):
            email = cellContent.value
            
    email = emailValidation(firstName, lastName, email)
    sheet["C" + str(row)] = email
    workBook.save(fileName)
    


workBook.close()




    
    

