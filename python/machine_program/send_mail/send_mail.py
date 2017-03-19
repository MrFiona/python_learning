#!/usr/bin/env python
# -*- coding: utf-8 -*-
# User    : apple
# File    : send_mail.py
# Author  : MrFiona 一枚程序员
# Time     : 2017-03-19 22:45

#hyyqgyyxxsughcgi

#smtp.qq.com


import smtplib
from email.mime.text import MIMEText
_user = "写自己的qq邮箱"
_pwd  = "这里写授权码" #授权码，在qq邮箱设置的开启IMAP/SMTP服务
_to   = "目的邮箱地址"

msg = MIMEText("Test")
msg["Subject"] = "我用来测试的邮箱轰炸器!!!就20封而已!!!"
msg["From"]    = _user
msg["To"]      = _to

send_num = 5
while send_num > 0:
    try:
        s = smtplib.SMTP_SSL("smtp.qq.com",465)
        s.login(_user, _pwd)
        s.sendmail(_user, _to, msg.as_string())
        s.quit()
        print "Success!"
    except smtplib.SMTPException,e:
        print "Falied,%s"%e
    send_num -= 1
