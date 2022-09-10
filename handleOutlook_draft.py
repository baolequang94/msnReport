import win32com.client as win32

from constant import OUTLOOK_APPLICATION, EXCEL_APPLICATION
from config import EXCEL_PATH,EMAIL_FROM, MORNING_SHIFT, AFTERNOON_SHIFT, EMAIL_TO, AFTERNOON_SHIFT_PERIODS
from helper import getSubject
from renderHtml import TABLE

outlook = win32.Dispatch(OUTLOOK_APPLICATION)
excel = win32.Dispatch(EXCEL_APPLICATION)

excel.Visible = False
excel.DisplayAlerts = False
excelApp = excel.Workbooks

trackerFile = excelApp.Open(EXCEL_PATH)
infopaneSheet = trackerFile.Worksheets("Infopane")

mail = outlook.CreateItem(0)
mail.Subject = getSubject(AFTERNOON_SHIFT)
mail.HTMLBody = TABLE

for index in range (len(AFTERNOON_SHIFT_PERIODS) - 1):
  
    output = open("tableRowTemplate.html", 'r', encoding='utf-8').read()
    output = output.replace(r"{%START%}", AFTERNOON_SHIFT_PERIODS[index])
    output = output.replace(r"{%END%}", AFTERNOON_SHIFT_PERIODS[index+1])
    print(output)
    

mail.To = EMAIL_FROM
# mail.CC = EMAIL_TO
# mail.Display()
# mail.Save()
# mail.Send()


# Alternative way to send email
# olNS = outlook.Session
# mail._oleobj_.Invoke(
#     *(64209, 0, 8, 0, olNS.Accounts.Item(EMAIL_FROM)))