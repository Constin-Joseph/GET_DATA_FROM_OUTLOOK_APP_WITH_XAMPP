import sys
sys.path.append('C:\\Program Files (x86)\\Python37-32\\Lib\\site-packages\\win32')
sys.path.append('C:\\Program Files (x86)\\Python37-32\\Lib\\site-packages\\win32\\lib')
import win32com.client
import os
import pymysql
import re
connection = pymysql.connect(
    host='localhost',
    user='root',
    password='',
    db='joseph',
)

outlook=win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
inbox=outlook.GetDefaultFolder(6) #Inbox default index value is 6
message=inbox.Items
message2=message.GetLast()
subject=message2.Subject
   
body=message2.body
date=message2.senton.date()
time=message2.senton.time() 
sender=message2.Sender
Cc=message2.Cc

attachments=message2.Attachments
print(subject)
text = []
print(body)
text = body.split('\n')
for t in re.split('\s+', subject):
    if re.findall("[a-z][-][0-9]", t):
        print(t)
    else:
         for i in text:
             if i.startswith("MS ID:"):
                txt1 = i.split(':')
                t=(txt1[1])        


for i in text:
     if i.startswith("MS TITLE:"):
        txt2 = i.split(':')
        MSTITLE=(txt2[1])
     else:
        MSTITLE=""
     if i.startswith("SPECIAL ISSUE:"):
        txt3 = i.split(':')
        si=(txt3[1])
     else:
        si=""
     if i.startswith("SPECIAL ISSUE NAME:"):
        txt4 = i.split(':')
        sin=(txt4[1])
     else:
        sin=""
     if i.startswith("Authors:"):
        txt5 = i.split(':')
        authors=(txt5[1])
     else:
        authors=""
     if i.startswith("Number of:"):
        txt6 = i.split(':')
        numberof=(txt6[1])
     else:
        numberof=""
     if i.startswith("APCs:"):
        txt7 = i.split(':')
        apcs=(txt7[1])
     else:
        apcs=""
     if i.startswith("LINKED PAPERS:"):
        txt8 = i.split(':')
        lp=(txt8[1])
     else:
        lp=""
print(sender)
print(attachments.count)
print(date)
print(time)
print(Cc)

try:
    with connection.cursor() as cursor:
        sql=("""INSERT INTO jack (subject,body,sender,date,time,MS_ID,Cc,MS_TITLE,SPECIAL_ISSUE_NAME,Authors,Number_of,APCs,LINKED_PAPERS,SPECIAL_ISSUE) VALUES ("%s","%s","%s","%s","%s","%s","%s","%s","%s","%s","%s","%s","%s","%s")"""%(subject,body,sender,date,time,t,Cc,MSTITLE,sin,authors,numberof,apcs,lp,si))
        try:
            cursor.execute(sql)
            print("Task added successfully")
        except:
            print("Oops! Something wrong")
 
    connection.commit()
finally:
    connection.close()

