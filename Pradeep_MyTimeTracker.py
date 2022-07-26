import pandas as pd
from datetime import date, datetime
import pytz
import win32com.client as win32

teamLeadsMail = ""

managerMail = ""

needToCheckOn=["Hours charged on MyTime","PTO/Sick leaves"]
EmailIDColumnName='Email ID'
today = date.today()
tz_NY = pytz.timezone('Asia/Kolkata')
datetime_NY = datetime.now(tz_NY)

todayDate = today.strftime("%B %d, %Y")
print(todayDate)

ccMails = ""

def send_mail_to_pending_ppl(val):

    if val is not None:
        toMails = ';'.join(val)
        try:
            send_Mail(toMails=toMails, ccMails=ccMails, date=todayDate)
        except Exception as inst:
            print("Some thing went wrong"+inst.with_traceback())

def send_Mail(toMails, ccMails, date):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = toMails
    mail.Subject = 'Gentle reminder to enter '+date
    mail.HTMLBody = '<h2>Tracker is not Updated</h2>'
    mail.CC = ccMails
    mail.Send()
    print("mailsent" + toMails)

def create_CCMails():
    hour = datetime_NY.hour
    mails=""
    if managerMail is not None and teamLeadsMail is not None:
        if(hour == 19):
            mails=teamLeadsMail
        if(hour ==20):
            mails=teamLeadsMail+";"+managerMail    
    return mails
def getMails(trackerName):
    df = pd.read_excel('C:\proj\MyTimeAutomation\MyTim.xlsx', sheet_name=trackerName)
    val=[]
    if df is not None:
        bool_series =pd.isna(df[needToCheckOn])
        df=df[bool_series]
        val=df.get(EmailIDColumnName).values.tolist()
    return val

if __name__ == "__main__":
    trackerName='tracker'
    ccMails=create_CCMails()
    val=getMails(trackerName)
    send_mail_to_pending_ppl(val)
