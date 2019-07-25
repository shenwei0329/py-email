#_*_coding:utf-8_*_
#

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
import time
import sys


class EmailClass(object):

    def __init__(self, info):
        self.curDateTime = str(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime()))
        self.sender = info['Smtp_Sender']
        self.receivers = info['Receivers']
        self.msg_title = info['Msg_Title']
        self.sender_server = info['Smtp_Server']
        self.passwd = info['Smtp_Password']
        self.From = info['From']
        self.To = info['To']
        self.Text = info['Text']

    def setMailContent(self):
        print self.receivers
        msg = MIMEMultipart()
        msg['From'] = self.From
        msg['To'] = self.To
        msg['Subject'] = Header('%s%s' % (self.msg_title, self.curDateTime), 'utf-8')
        msg.attach(MIMEText(self.Text, 'plain', 'utf-8'))
        return msg

    def addAttach(self,apath,filename='Report.html'):
        with open(apath, 'rb') as fp:
            attach = MIMEBase('application', 'octet-stream')
            attach.set_payload(fp.read())
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            encoders.encode_base64(attach)
            fp.close()
            return attach

    def sendEmail(self, message):
        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(self.sender_server, 25)
            smtpObj.login(self.sender, self.passwd)
            smtpObj.sendmail(self.sender, self.receivers, message.as_string())
            smtpObj.quit()
            print u"邮件发送成功"
        except smtplib.SMTPException as ex:
            print u"Error: 无法发送邮件.%s"% ex

    def send(self):
        self.sendEmail(self.setMailContent())


if __name__ == "__main__":

    mail = {
        "Smtp_Server": "smtp.chinacloud.com.cn",
        "Smtp_Password": sys.argv[1],
        "Receivers": ["shenwei0329@hotmail.com", "shenwei@chinacloud.com.cn"],
        "Msg_Title": "An Auto-Reply email by R&D MIS",
        "Smtp_Sender": "shenwei@chinacloud.com.cn",
        "From": "shenwei_from@chinacloud.com.cn",
        "To": "shenwei_to@chinacloud.com.cn",
        "Text": "For test it!"
    }
    EmailClass(mail).send()


