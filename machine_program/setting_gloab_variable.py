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

#public_use_function.py
CONFIG_FILE_PATH = os.getcwd() + os.sep + 'machineConfig' + os.sep + 'email.conf'

type_sheet_name_list = ['Save-Miss', 'NewSi', 'ExistingSi', 'CaseResult']
type_sheet_name_list.sort(reverse=True)