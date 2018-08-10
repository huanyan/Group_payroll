# -*- coding: UTF-8 -*-
# 调用库
import os
import xlwings as xw
import urllib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.utils import formatdate
from email import encoders
import shutil
import time
# 获取本地时间
localtime = time.asctime(time.localtime(time.time()))
# 创建员工工资表文件夹
path = os.path.join(os.getcwd(), 'send_mail', time.strftime("%Y-%m-%d_%H%M%S", time.localtime()))
os.makedirs(path)
# 工资条类
# 邮箱字典
wb_emial = xw.Book(r'员工_邮箱.xlsx')
email = wb_emial.sheets[0].range('B2').expand('down').value
email_name = wb_emial.sheets[0].range('A2').expand('down').value
email_dic = dict(zip(email_name, email))
wb_emial.close()
# 员工名单数组
wb_money = xw.Book(r'工资条.xlsx')
name = wb_money.sheets[0].range('C3').expand('down').value
name_length = range(len(name))
for i in name_length:
    x = i + 3
    new_name = str(name[i]) + '.xlsx'
    new_book = xw.Book()
    new_book.sheets[0].range('A1').value = wb_money.sheets[0].range('B2:C2').value + wb_money.sheets[0].range('K2:L2').value + wb_money.sheets[0].range('T2:AA2').value + wb_money.sheets[0].range('AD2:AE2').value + wb_money.sheets[0].range('AG2:AH2').value + wb_money.sheets[0].range('AK2:AL2').value
    new_book.sheets[0].range('A2').value = wb_money.sheets[0].range((x, 2), (x, 3)).value + wb_money.sheets[0].range((x, 11), (x, 12)).value + wb_money.sheets[0].range((x, 20), (x, 27)).value + wb_money.sheets[0].range((x, 30), (x, 31)).value + wb_money.sheets[0].range((x, 33), (x, 34)).value + wb_money.sheets[0].range((x, 37), (x, 38)).value
    new_book.save(os.path.join(path, new_name))
    new_book.save()
    new_book.close()
wb_money.close()
# 邮件群发
succese_numbers = 0
senderMail = input("请输入发件邮箱:")
senderPass = input("请输入邮箱密码:")
for key in email_dic:
    mail_host = 'smtp.163.com'
    mail_user = str(senderMail[:-8])
    mail_pass = str(senderPass)
    sender = str(senderMail)
    receivers = str(email_dic[key])
    # 设置email信息
    message = MIMEMultipart()
    message['Subject'] = '工资条'
    message['From'] = sender
    message['To'] = receivers
    message['Date'] = formatdate(localtime=True)
    att = MIMEApplication(open('send_mail/'+path[-17:] + '/' + key + '.xlsx', 'rb').read())
    att['Content-Type'] = 'application/octet-stream'
    att['Content-Disposition'] = 'attachment;filename="PaySlip.xlsx"'
    message.attach(att)
    # 登陆并发邮件
    try:
        smtp0bj = smtplib.SMTP()
        smtp0bj.connect(mail_host, 25)
        smtp0bj.login(mail_user, mail_pass)
        smtp0bj.sendmail(sender, receivers, message.as_string())
        smtp0bj.quit()
        print(receivers + '的工资单发送成功')
        succese_numbers = succese_numbers + 1
    except smtplib.SMTPException as e:
        print(receivers + '的工资单发送失败', e)
email_log = open("email_log.txt", 'r+')
if succese_numbers == len(email):
    email_log.read()
    email_log.write("\n所有工资单群发成功!" + "  " + "时间:" + localtime)
    email_log.close()
else:
    email_log.read()
    email_log.write("\n工资单群发失败!请检查输入邮箱账号和密码是否正确." + "  " + "时间:" + localtime + "\n")
    email_log.close()
os.system('pause')
