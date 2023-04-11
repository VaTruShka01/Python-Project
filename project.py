import openpyxl
import win32com.client as win32
import pyinputplus as pip

# asks user file path on their local computer
fileName = pip.inputFilepath(
    prompt="Please enter spreadsheet filename, ensure that there are no quotes. (spreadsheet should have columns: A - first name, B - last name, C - email)\n")

# defines workBook, sheet, list of columns and count of rows
workBook = openpyxl.load_workbook(fileName)
sheet = workBook.active
columnList = ["A", "B", "C"]
rowCount = sheet.max_row

# email validation
def emailValidation(firstName, lastName, email):
    firstName = str(firstName).lower().replace(" ", "")
    lastName = str(lastName).lower().replace(" ", "")

    basicPattern = firstName + "." + lastName + "@georgiancollege.ca"

    # creates email if the field is empty
    if (email == None):
        email = basicPattern
        print("Email cell", str(column) + str(row) + ", of", firstName.title(),
              lastName.title(), "was empty so a new email was generated automatically:", email)

    # cleans the email from spaces
    email = email.replace(" ", "").lower()

    # checks if email is the same as required pattern
    if (email != basicPattern and email != None):
        email = basicPattern
        print("Email cell", str(column) + str(row) + ", of", firstName.title(),
              lastName.title(), "has invalid format and was changed to:", email)

    return email

# iterates through all cells in spreadsheet
for row in range(2, rowCount + 1):
    for column in columnList:
        cellContent = sheet[str(column)+str(row)]

        # checks and tells to user that cell is empty
        if (cellContent == None and column != "C"):
            print("Cell", str(column) + str(row), "is empty")
            # sheet.delete_row(row, 1)
            # workBook.save(fileName)

        # defines variables according to cell value
        if (column == "A"):
            firstName = cellContent.value
        elif (column == "B"):
            lastName = cellContent.value
        elif (column == "C"):
            email = cellContent.value

    # validates each email and reassigns value to cell
    email = emailValidation(firstName, lastName, email)
    sheet["C" + str(row)] = email
    workBook.save(fileName)


# checks all emails to send messages
for row in range(2, rowCount + 1):

    cellContent = sheet['C' + str(row)]
    cellContent = cellContent.value

    try:
        # creates object of outlook using application in PC
        outlook = win32.Dispatch('outlook.application')
    except:
        print("You do not have Outlook application installed on your PC")

    # creates mail object that will be sent
    mail = outlook.CreateItem(0)
    mail.To = cellContent
    mail.Subject = 'Welcome to GC Flex Teaching!'
    mail.HTMLBody = '''
    <!DOCTYPE html>
    <html>
        <body>
            <div class="x_x_x_elementToProof" style="font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)">Thank you for your hard work to become a GC Flex instructor. The team at CTL is here to support you as you transition to this new role. Here are some ways you can find support:</div>
            <div class="x_x_x_elementToProof" style="font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)"><br aria-hidden="true"></div>
            <div class="x_x_x_elementToProof" style="font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)"><ul><li class="x_x_x_elementToProof">Sign up for an in-person flex technology demo <span><a href="https://georgiancollege-my.sharepoint.com/:w:/g/personal/corey_berry_georgiancollege_ca/EY0wOLx-Xg9JsyYNz9OnirsBL7tJEgs1HtGSe3x1wbaETA?e=bWeS3k" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" id="OWA3a4dac0e-fe5e-ec19-fd9e-fed254c78b05" class="x_x_x_OWAAutoLink x_x_x_WSYlv" data-ogsc="" data-loopstyle="link" data-safelink="true" data-linkindex="0">using this form</a></span></li><li class="x_x_x_elementToProof">Attend a virtual meeting about MS Teams and what to do before your first class (see calendar invite)</li><li class="x_x_x_elementToProof">Check out the <a href="https://teams.microsoft.com/l/channel/19%3a1_CvxzXFL27CmadTBZgukQTI0HOaFaEjUoKiYFbdL3s1%40thread.tacv2/General?groupId=6dc596da-aeb3-4cf3-8f97-27ff0db4d8de&amp;tenantId=da9a94b6-4681-49bc-bd7c-bab9eac0ad3c" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" title="https://teams.microsoft.com/l/channel/19%3a1_CvxzXFL27CmadTBZgukQTI0HOaFaEjUoKiYFbdL3s1%40thread.tacv2/General?groupId=6dc596da-aeb3-4cf3-8f97-27ff0db4d8de&amp;tenantId=da9a94b6-4681-49bc-bd7c-bab9eac0ad3c" data-loopstyle="link" id="LPNoLPOWALinkPreview" data-safelink="true" data-linkindex="1">GC Flex Support Team</a> and engage with other instructors that have taught in GC Flex</li><li class="x_x_x_elementToProof">Attend virtual drop-in office hours (see your calendar for invite)&nbsp;</li><li class="x_x_x_elementToProof">Email facultyedtech@georgiancollege.ca with any questions or concerns</li></ul><div class="x_x_x_elementToProof">Wishing you all the best for a successful term!</div><div class="x_x_x_elementToProof"><br aria-hidden="true"></div><div class="x_x_x_elementToProof">Sincerely,</div><div class="x_x_x_elementToProof">Your GC Flex Support Team</div><div class="x_x_x_elementToProof"><br aria-hidden="true"></div><div class="x_x_x_elementToProof"><b style="font-style:inherit; font-variant-ligatures:inherit; font-variant-caps:inherit; font-family:&quot;Segoe UI&quot;,sans-serif; font-size:13.3333px; color:rgb(32,31,30)"><span style="color:rgb(31,78,121)">Corey Berry, OCT, MA&nbsp;</span></b><span style="font-weight:400; font-size:13.3333px; font-family:&quot;Segoe UI&quot;,sans-serif,serif,EmojiFont; color:rgb(31,78,121)">(she/her)</span><br aria-hidden="true"></div></div>
            <div class="x_x_x_elementToProof"><div id="x_x_x_Signature"><div><div style="font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont; font-size:12pt; color:rgb(0,0,0)"><b style="font-family:&quot;Segoe UI&quot;,sans-serif; font-size:13.3333px; color:rgb(32,31,30)"><span style="margin:0px; color:rgb(112,173,71)">Instructional Design Technologist, Centre for Teaching &amp; Learning (CTL)</span></b> <p style="margin:0in 0in 0.0001pt; font-size:11pt; font-family:Calibri,sans-serif; text-align:start; color:rgb(32,31,30); background-color:rgb(255,255,255)"><b><span style="margin:0px; color:black"></span></b></p><p style="margin:0in 0in 0.0001pt; font-size:11pt; font-family:Calibri,sans-serif; text-align:start; color:rgb(32,31,30); background-color:rgb(255,255,255)"><span style="margin:0px; color:black">Georgian College| One Georgian Drive | Barrie&nbsp;ON |&nbsp;L4M&nbsp;3X9<br aria-hidden="true"><br aria-hidden="true"></span></p><div style="margin:0px 0in 0.000133333px; font-size:11pt; font-family:Calibri,sans-serif,serif,EmojiFont; text-align:start; color:rgb(32,31,30); background-color:rgb(255,255,255)"><span data-contrast="auto" lang="EN-US" style="margin:0px; text-align:left; font-size:11pt; line-height:19.425px; font-family:Calibri,Calibri_EmbeddedFont,Calibri_MSFontService,sans-serif,serif,EmojiFont; color:rgb(0,0,0); background-color:rgb(255,255,255); font-variant-ligatures:none!important"><span style="margin:0px"><b>Did you know</b> that you can request your own educational technology workshops</span><span style="margin:0px">? </span><span style="margin:0px">Use this </span></span><a href="https://forms.office.com/Pages/ResponsePage.aspx?id=tpSa2oFGvEm9fLq56sCtPNb7byXSFtBKqzrjOGp6I89UNFFRT1BTRjVDRkQ4VlJVQ09FNkszRzNQSy4u" target="_blank" rel="noreferrer noopener" data-auth="NotApplicable" data-safelink="true" data-linkindex="2" style="margin:0px; font-family:&quot;Segoe UI&quot;,&quot;Segoe UI Web&quot;,Arial,Verdana,sans-serif; font-size:12px; text-align:left; background-color:rgb(255,255,255)"><span data-contrast="none" lang="EN-US" style="margin:0px; font-size:11pt; text-decoration:underline; line-height:19.425px; font-family:Calibri,Calibri_EmbeddedFont,Calibri_MSFontService,sans-serif,serif,EmojiFont; color:rgb(5,99,193); font-variant-ligatures:none!important"><span data-ccp-charstyle="Hyperlink" style="margin:0px">form to request a workshop</span></span></a>.<br aria-hidden="true"></div><p style="margin:0in 0in 0.0001pt; font-size:11pt; font-family:Calibri,sans-serif; text-align:start; color:rgb(32,31,30); background-color:rgb(255,255,255)"><br aria-hidden="true"></p><div style="margin:0px 0in 0.000133333px; font-size:11pt; font-family:Calibri,sans-serif,serif,EmojiFont; text-align:start; color:rgb(32,31,30); background-color:rgb(255,255,255)"><p style="font-size:10pt; font-family:&quot;Segoe UI&quot;,sans-serif; margin:0px; background-color:rgb(255,255,255)"><i>Part of my of Truth and Reconciliation journey is to acknowledge that the land I live, work, and play on are the traditional lands of the Anishinaabeg, Haudenosaunee, Tionontati, and Wendat people. I strive to be an ally to Indigenous people by '<a href="https://ecampusontario.pressbooks.pub/skoden/chapter/living-together-in-a-good-way/" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" title="https://ecampusontario.pressbooks.pub/skoden/chapter/living-together-in-a-good-way/" data-safelink="true" data-linkindex="3">living together in a good way</a>', learning, and listening.&nbsp;</i></p></div><br aria-hidden="true"></div></div></div></div>
        </body>
    </html> 
    '''
    mail.Send()

workBook.close()