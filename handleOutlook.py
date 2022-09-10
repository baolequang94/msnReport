import sys
import win32com.client as win32

from constant import OUTLOOK_APPLICATION
from config import EMAIL_TO, EMAIL_CC
from helper import getSubject

def handleOutlook(shift, attachmentPath):
    print("Opening Outlook")
    outlook = win32.Dispatch(OUTLOOK_APPLICATION)
    mail = outlook.CreateItem(0)
    mail.Subject = getSubject(shift)
    mail.To = ";".join(EMAIL_TO)
    mail.CC = ";".join(EMAIL_CC)
    mail.Display()
    mail.Attachments.Add(attachmentPath)
    mail.Save()

sys.modules[__name__] = handleOutlook

# mail.HTMLBody = TABLE
# mail.Send()

# Alternative way to send email
# olNS = outlook.Session
# mail._oleobj_.Invoke(
#     *(64209, 0, 8, 0, olNS.Accounts.Item(EMAIL_FROM)))


# emailDraft = "lequangbao.l@hcl.com"
# emailDraft2 = "v-nguyenhoan@microsoft.com"