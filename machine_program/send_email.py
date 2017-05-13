#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-06 10:43
# Author  : MrFiona
# File    : send_email.py
# Software: PyCharm Community Edition


import os
import smtplib
import glob
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from email.mime.multipart import MIMEMultipart
from machine_config import MachineConfig
from setting_gloab_variable import DEBUG_FLAG, PURL_BAK_STRING
from public_use_function import deal_html_data, get_url_list_by_keyword


class SendEmail:
    def __init__(self):
        self.config_file_path = os.getcwd() + os.sep + 'machineConfig' + os.sep + 'email.conf'
        deal_html_data()
        self.send()

    def send_mail(self):
        # 读取配置文件日志、时间设置
        conf = MachineConfig(self.config_file_path)
        SERVER = conf.get_node_info('server', 'server_address')
        FROM = conf.get_node_info('from_address', 'address')
        TO = conf.get_node_info('receive_address', 'address')
        SUBJECT = conf.get_node_info('mail_format', 'subject')
        TEXT = conf.get_node_info('mail_format', 'text')
        FILE = conf.get_node_info('mail_format', 'file')
        IMAGE = conf.get_node_info('mail_format', 'image')
        HTML = conf.get_node_info('mail_format', 'html')

        if DEBUG_FLAG:
            print '\033[31mfrom address:\033[0m \033[32m%s\033[0m' % FROM
            print '\033[31mto address:\033[0m \033[32m %s\033[0m' % TO
            print '\033[31mserver address:\033[0m \033[32m %s\033[0m' % SERVER
            print "\033[36m********** Message body **********\033[0m"
            print "         \033[37msubject: %s\033[0m" % SUBJECT
            print "         \033[37mtext: %s\033[0m" % TEXT
            print "         \033[37mtext: %s\033[0m" % FILE
            print "         \033[37mtext: %s\033[0m" % IMAGE
            print "         \033[37mtext: %s\033[0m" % HTML
            print "\033[36m**********************************\033[0m"

        message = """From: %s\r\nTo: %s\r\nSubject: %s\r\n\

          %s
          """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

        server = smtplib.SMTP(SERVER)
        server.sendmail(FROM, TO, message)

    def _format_addr(self, s):
        name, addr = parseaddr(s)
        return formataddr((
            Header(name, 'utf-8').encode(),
            addr.encode('utf-8') if isinstance(addr, unicode) else addr))

    def send(self):
        config_file_path = os.getcwd() + os.sep + 'machineConfig' + os.sep + 'email.conf'
        conf = MachineConfig(config_file_path)
        from_addr = conf.get_node_info('from_address', 'address')
        to_addr = conf.get_node_info('receive_address', 'address')
        smtp_server = conf.get_node_info('server', 'server_address')

        Silver_url_list = get_url_list_by_keyword(PURL_BAK_STRING, 'Silver')
        lastest_week_string = Silver_url_list[0]
        lastest_week_string = 'WW' + lastest_week_string.split('/')[-2].split('%')[-1].split('WW')[-1]

        SUBJECT = u'BKC ITF Test Plan of Purley-FPGA %s……' % lastest_week_string

        if DEBUG_FLAG:
            print '\033[31mfrom address:\033[0m \033[32m%s\033[0m' % from_addr
            print '\033[31mto address:\033[0m \033[32m%s\033[0m' % to_addr
            print '\033[31mserver address:\033[0m \033[32m %s\033[0m' % smtp_server
            print "\033[36m***************************** Message body *****************************\033[0m"
            print "         \033[37msubject: %s\033[0m" % SUBJECT
            print "\033[36m************************************************************************\033[0m"

        html_string = ''

        if os.path.exists(os.getcwd() + os.sep + 'html_result'):
            for file in glob.glob(os.getcwd() + os.sep + 'html_result' + os.sep + '*.html'):
                with open(file, 'r') as f:
                    html_data = f.readlines()
                for i in range(len(html_data)):
                    html_data[i] = html_data[i].strip()
                    html_string += html_data[i]
        # 邮件对象:
        # msg = MIMEMultipart()
        # 邮件正文是MIMEText:
        # msg.attach(MIMEText('This is a test message written by python...', 'plain', 'utf-8'))
        msg = MIMEText(html_string, 'html', 'utf-8')
        msg['From'] = self._format_addr(from_addr)
        # 添加附件就是加上一个MIMEBase
        # with open(r'C:\Users\pengzh5x\PycharmProjects\personal_program\machine_program\excel_dir\report_result.xlsx', 'rb') as f:
        #     # 设置附件的MIME和文件名，这里是png类型:
        #     mime = MIMEBase('file', 'xlsx', filename='report_result.xlsx')
        #     # 加上必要的头信息:
        #     mime.add_header('Content-Disposition', 'attachment', filename='report_result.xlsx')
        #     mime.add_header('Content-ID', '<0>')
        #     mime.add_header('X-Attachment-Id', '0')
        #     # 把附件的内容读进来:
        #     mime.set_payload(f.read())
        #     # 用Base64编码:
        #     encoders.encode_base64(mime)
        #     # 添加到MIMEMultipart:
        #     msg.attach(mime)

        address_list = []
        for address in to_addr.split(','):
            if len(address) == 0:
                continue
            address_list.append(address)

        msg['To'] = ','.join(address_list)
        msg['Subject'] = Header(u'BKC ITF Test Plan of Purley-FPGA %s……' % lastest_week_string, 'utf-8').encode()
        server = smtplib.SMTP(smtp_server, 25)
        server.set_debuglevel(1)
        server.sendmail(from_addr, address_list, msg.as_string())
        server.quit()

if __name__ == '__main__':
    send = SendEmail()
