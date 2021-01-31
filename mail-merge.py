import pandas
import smtplib,os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
class emailsend:
  def __init__(self,last_name,salutation,Email_Address):
        self.ln=last_name
        self.salut=salutation
        self.Email=Email_Address
filename="C:\\Users\\Shiv Chandra Sraswat\\Desktop\\mail.xlsx"
def readexcel(filename):
    x=pandas.read_excel(filename)
    
    flag=0
    thread=[]
    for i in range(len(x)):
       node=emailsend(last_name=x.iloc[i,1],salutation=x.iloc[i,2],Email_Address=x.iloc[i,3])
       thread.append(node)
    return thread    
server=input("Enter your mail type eg gmail,outlook etc")
user=input("Enter your username/email :")
pwd=input("Enter mail password :")
subject=input("Enter subject")
cc_address=input("Enter cc address")
if server=="Gmail":
   pt="smtp.gmail.com"
elif (server=="Outlook"or server=="Live"or server=="outlook"):
   pt="smtp-mail.outlook.com"
thread=readexcel(filename)
t=len(thread)
x=smtplib.SMTP(pt,587)
x.ehlo()
x.starttls()
x.login(user,pwd)

for i in thread:
        b=i.Email
        a=i.ln
        salut=i.salut
        
        strFrom =user
        strTo =b
        msgRoot = MIMEMultipart('related')
        msgRoot['Subject'] = subject
        msgRoot['From'] = strFrom
        msgRoot['To'] = strTo
        msgRoot['Cc']=cc_address
        msgRoot.preamble = 'This is a multi-part message in MIME format.'
        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)
        msgText = MIMEText('hi')
        msgAlternative.attach(msgText)
        c=''' enter text here '''
        
        msgText = MIMEText(c, 'html')
        msgAlternative.attach(msgText)
        x.sendmail(strFrom, strTo,msgRoot.as_string())  
print("email sent")
print("the email files are not kept now")
print("Thank you for trusting our app")

