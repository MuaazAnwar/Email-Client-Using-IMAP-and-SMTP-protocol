from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui, QtWidgets
import imaplib, email, os,sys,datetime
import smtplib,time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import easygui
import ntpath

app = QtWidgets.QApplication([])

pg1 = uic.loadUi("login.ui")
pg2 = uic.loadUi("Received Mail.ui")
pg3=uic.loadUi("sendmail.ui")

con = imaplib.IMAP4_SSL('imap.gmail.com')
def loginpressed():
    
    user= pg1.TUname.text()
    password=pg1.Tpass.text()
    #imap_url = 'imap.gmail.com'
    global con
    con = imaplib.IMAP4_SSL('imap.gmail.com')
    try:
        raw,data=con.login(user,password)
        if raw=='OK':
            pg1.hide()
            pg2.show()
            res,data=con.select('INBOX')
            temp = str(int(data[0]))
            pg2.Ttotmail.setText(temp)
            
            
    except:
        print('Unexpected error : {0}'.format(sys.exc_info()[0]))
        pg1.Lerror.setText("Invalid Username or password") 
    
    return con
def setfolder(type):
    try:
        global con
        res,data=con.select(type)
        temp = str(int(data[0]))
        pg2.Ttotmail.setText(temp)
    except:
        print('Unexpected error : {0}'.format(sys.exc_info()[0]))
        print("i am not working")
def showmail():
    try:
        global con

        result, data = con.fetch(pg2.Tinputmail.text(),'(RFC822)')
        global email_message 
        email_message = email.message_from_bytes(data[0][1])
        date_tuple = email.utils.parsedate_tz(email_message['Date'])
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
            local_message_date= "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
    
        email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
        email_to = str(email.header.make_header(email.header.decode_header(email_message['To'])))
        subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
        pg2.Theader.setText("From: %s\nTo: %s\nDate: %s\nSubject: %s\n " %(email_from, email_to,local_message_date, subject))
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body= part.get_payload(decode=True)
                pg2.Tbody.setText(body.decode('utf-8'))
        
        pg2.Lerror.setText("")

    except:
        pg2.Lerror.setStyleSheet("QLabel {color:red;}")
        pg2.Lerror.setText("Invalid Number!")

def attachmentdownload():
    ###/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    attachment_dir=easygui.diropenbox()
    if not attachment_dir:
    	return
    

    pg2.Tpath.setText(attachment_dir)
    attachment_dir = os.path.abspath(attachment_dir)
    #print(os.path.abspath(attachment_dir))
    ##/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


    global email_message
    flag=0
    try:
        for part in email_message.walk():
            if part.get_content_maintype()=='multipart':
                 
                continue
            if part.get('Content-Disposition') is None:
                
                continue
                
                 
            fileName = part.get_filename()
                
            if bool(fileName):
                flag=1
                filePath = os.path.join(attachment_dir, fileName)
                with open(filePath,'wb') as f:
                    f.write(part.get_payload(decode=True))

        if flag==0 :
            pg2.Lerror2.setStyleSheet("QLabel {color:red;}")
            pg2.Lerror2.setText("No attachment found!")


        else:
            pg2.Lerror2.setText("Download Completed")
    except:
        pg2.Lerror2.setStyleSheet("QLabel {color:red;}")
        pg2.Lerror2.setText("No attachment found!")

def sentmail():
    pg3.show()

def nextpressed():
    pg2.Tinputmail.setText(str(int(pg2.Tinputmail.text())+1))
    showmail()
def prepressed():
    pg2.Tinputmail.setText(str(int(pg2.Tinputmail.text())-1))
    showmail()

def sendmailsmtp():
    if pg3.smail.toPlainText():
        fromaddr=pg1.TUname.text()
        toaddr = pg3.smail.toPlainText()
        msgg = pg3.msg.toPlainText()
        subj = pg3.sub.toPlainText()

        message = MIMEMultipart()
        message['From'] = fromaddr
        message['To'] = toaddr
        message['Subject'] = subj

        message.attach(MIMEText(msgg, 'plain'))
        if pg3.sub_2.toPlainText():
            attach_file = open(pg3.attach_file_name, 'rb')
            payload = MIMEBase('application', 'octate-stream')
            payload.set_payload((attach_file).read())
            encoders.encode_base64(payload)

            payload.add_header('Content-Disposition','attachment', filename=ntpath.basename(pg3.attach_file_name))
            message.attach(payload)
        try:
            server = smtplib.SMTP("smtp.gmail.com:587")
            server.set_debuglevel(1)
            server.ehlo()
            server.starttls()
            server.login(fromaddr, pg1.Tpass.text())
            text = message.as_string()
            ret = server.sendmail(fromaddr, toaddr, text)

        except Exception as e:
                print('some error occured')
                print(e)
                pg3.Linfo.setText(" ")
                pg3.Lerror2.setStyleSheet("QLabel {color:red;}")
                pg3.Lerror2.setText("Invalid Sender!")
        else:
                pg3.Linfo.setText("Mail Sent Sucessfully!")
                pg3.Lerror2.setText(" ")
                #time.sleep(2)
                #pg3.Linfo.setText(" ")

    else:
        pg3.Linfo.setText(" ")
        pg3.Lerror2.setStyleSheet("QLabel {color:red;}")
        pg3.Lerror2.setText("No sender email id is found!")

def setpath():
    a = easygui.fileopenbox()
    if a:
        pg3.sub_2.setText(a)
        pg3.attach_file_name = os.path.abspath(a)


if __name__=="__main__":
    pg1.LPic.setPixmap(QtGui.QPixmap("img/email.png"))
    pg1.LPic.setScaledContents(True)
    pg1.Lerror.setStyleSheet("QLabel {color:red;}")
    con=pg1.loginButton.clicked.connect(loginpressed)
    pg2.showButton.clicked.connect(showmail)
    pg2.downloadButton.clicked.connect(attachmentdownload)
    pg2.sentButton.clicked.connect(lambda: setfolder('"[Gmail]/Sent Mail"'))
    pg2.inboxButton.clicked.connect(lambda: setfolder('INBOX'))
    pg2.spamButton.clicked.connect(lambda: setfolder('[Gmail]/Spam'))
    pg2.trashButton.clicked.connect(lambda: setfolder('[Gmail]/Trash'))
    pg2.composeButton.clicked.connect(sentmail)
    pg2.nextButton.clicked.connect(nextpressed)
    pg2.preButton.clicked.connect(prepressed)
    pg3.sendbtn.clicked.connect(sendmailsmtp)
    pg3.browsebtn.clicked.connect(setpath)

    pg1.show()
    app.exec()
