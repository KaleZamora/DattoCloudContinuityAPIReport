import SendEmail
import json
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32
import tzlocal
import time

#Important note to those who come after me, this thing throws alot of warnings from the misuse of dataframe manipulation
#However it is the most efficent way to get from point A to point B.
#You should also be aware that Nucleus is the only catch statement, if other sites use illegal chars it will break the
#script. There is also a try except statement thats only job is to ignore empty sites in Datto (site created, no agents)
#Removing that will break the script. When you inherit this, change your email for mine or the next guys email so it will
#Alert you when it breaks. -ZK



try:
    now = time.time()
    parameters = {"_perPage": 100}
    response = requests.get(r"https://api.datto.com/v1/bcdr/agent", auth=HTTPBasicAuth("" , ""))
    text = response.text
    data = json.loads(text)
    data = data['clients']
    for i in range(len(data)):
        clientname = data[i]
        #print(clientname)
        clientname2 = clientname['clientName']
        #print(clientname2)
        agents = clientname['agents']
        #print(agents)
        df = pd.DataFrame()
        for i in agents:
            df = pd.DataFrame(agents)
            #print(df['lastSnapshot'])
        try:
            runner = range(len(df['lastSnapshot']))
        except:
            print(clientname2)
            continue
        fixer = df.copy()
        switcher = fixer['lastSnapshot']
        switcher2 = df.copy()
        for i in runner:
            test = datetime.fromtimestamp(switcher2['lastSnapshot'][i])
            switcher2['lastSnapshot'][i] = test
        switcher2.drop(labels='uuid', axis=1, inplace=True)
        switcher2.drop(labels='shortCode', axis=1, inplace=True)
        switcher2.drop(labels='type', axis=1, inplace=True)
        switcher2.drop(labels='lastScreenshot', axis=1, inplace=True)
        switcher2.drop(labels='screenshotSuccess', axis=1, inplace=True)
        switcher2.drop(labels='protectedMachine', axis=1, inplace=True)

        today = datetime.now()
        check = today - timedelta(days=14)
        finalresult = switcher2[switcher2['lastSnapshot'] <= check]
        check = check.strftime('%m-%d-%Y')
        today = today.strftime('%m-%d-%Y')
        #Sprint(finalresult)
        try:
            finalresult.loc[-1] = ["1969-12-31", "Means never backed up"]
        except:
            continue
        bad = '|'
        if bad in clientname2:
            clientname2 = 'Nucleus'
        finalresult2 = 'c:\\util\\' + clientname2 + '-olderthan14days-' + str(today) + ".xlsx"
        finalresult.to_excel('c:\\util\\' + clientname2 + '-olderthan14days-' + str(today) + ".xlsx")
        SendEmail.sendmail(today, finalresult2, clientname2)
except:

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = "You"
    #mail.To = ""
    mail.Subject = "Automate DCC Reporting Error"
    mail.body = "Shit broke mate."
    #mail.Attachments.Add(finalresult2)
    #mail.CC =
    mail.Send()
