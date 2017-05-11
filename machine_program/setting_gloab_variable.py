#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-05-09 10:57
# Author  : MrFiona
# File    : setting_gloab_variable.py
# Software: PyCharm Community Edition

import os

#machine_model_entrance.py
#预测 Purley-FPGA 需要设置下面值
PURL_BAK_STRING = 'Purley-FPGA'
#预测 Bakerville 需要设置下面值
# PURL_BAK_STRING = 'Bakerville'

#public_use_function.py
DO_PROF = True
BACKUP_EXCEL_DIR = os.getcwd() + os.sep + 'backup_dir'
SRC_EXCEL_DIR = os.getcwd() + os.sep + 'excel_dir'

#send_email.py
DEBUG_FLAG=True


def get_url_list_by_keyword(pre_keyword, back_keyword, key_url_list=None):
    if not key_url_list:
        key_url_list = []

    f = open(os.getcwd() + os.sep + 'report_html' + os.sep + 'url_info.txt')
    for line in f:
        if pre_keyword in line and back_keyword in line:
            key_url_list.append(line.strip('\n'))
    # print key_url_list, len(key_url_list)
    return key_url_list

# url列表初始化
Silver_url_list = get_url_list_by_keyword(PURL_BAK_STRING, 'Silver')

type_sheet_name_list = ['Save-Miss', 'NewSi', 'ExistingSi', 'CaseResult']
type_sheet_name_list.sort(reverse=True)

lastest_week_string = Silver_url_list[0]
lastest_week_string = 'WW' + lastest_week_string.split('/')[-2].split('%')[-1].split('WW')[-1]