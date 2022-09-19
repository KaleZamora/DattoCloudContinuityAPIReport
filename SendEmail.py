import win32com.client as win32
import datetime

def sendmail(today, finalresult2, clientname2):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = ""
    #mail.To = ""
    mail.Subject = "Inconsistencies discovered at " + str(clientname2)
    mail.body = "Please note the following Inconsistencies detailed in the attachment."
    mail.Attachments.Add(finalresult2)
    #mail.CC =
    mail.Send()
