#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-05-09 09:40
# Author  : MrFiona
# File    : create_email_html.py
# Software: PyCharm Community Edition

import os
import sys
import time

try:
    from pyh import PyH, h1, table, td, tr

except ImportError:
    pwd_dir = os.getcwd()
    os.chdir(os.getcwd() + os.sep + 'PyH-0.1.1')
    os.system('python setup.py install')
    os.chdir(pwd_dir)
    from pyh import PyH, h1, table, td, tr

from public_use_function import get_report_data, get_url_list_by_keyword
from setting_gloab_variable import PURL_BAK_STRING

try:
  import cPickle as pickle
except ImportError:
  import pickle

reload(sys)
sys.setdefaultencoding('utf-8')


def create_save_miss_html(sheet_name, Silver_url_list, purl_bak_string='Purley-FPGA'):
    if sheet_name == 'Save-Miss':
        page = PyH('excel结果html表格数据')
        page << h1(sheet_name + ' result')
        cell_data_list = get_report_data(sheet_name, purl_bak_string=purl_bak_string)
        # print cell_data_list

        mytab = table(border="1", cellspacing="1px")
        for i in range(len(cell_data_list)):
            mytr = mytab << tr()
            for j in range(len(Silver_url_list) + 3):
                if j == 0:
                    mytr << td(cell_data_list[i][j], style="width:60px")
                else:
                    mytr << td(cell_data_list[i][j])
        page << mytab

    elif sheet_name in ('ExistingSi', 'NewSi'):
        page = PyH('excel结果html表格数据')
        page << h1(sheet_name + ' result')
        cell_data_list = get_report_data(sheet_name, purl_bak_string=purl_bak_string)
        mytab = table(border="1", cellspacing="1px")

        for k in range(len(cell_data_list)):
            cell_data_list[k] = cell_data_list[k][2:]

        for i in range(len(cell_data_list)):
            mytr = mytab << tr()
            for j in range(len(cell_data_list[0])):
                if j == 1:
                    mytr << td(cell_data_list[i][j], style="width:150px")
                else:
                    mytr << td(cell_data_list[i][j])
        page << mytab

    else:
        page = PyH('excel结果html表格数据')
        page << h1(sheet_name + ' result')
        cell_data_list = get_report_data(sheet_name, purl_bak_string=purl_bak_string)
        mytab = table(border="1", cellspacing="1px")

        cell_data_list[0][2] = ''
        for i in range(len(cell_data_list)):
            mytr = mytab << tr()
            for j in range(len(cell_data_list[0])):
                if i == 0 and j == 0:
                    mytr << td(cell_data_list[i][j], style="background-color:#D2691E;width:60px")
                elif i == 0 and j == 1:
                    mytr << td(cell_data_list[i][j], style="font-size:30px;color:#00008B;background-color:#FFDEAD;width:50px")
                elif i == 0 and j == 2:
                    mytr << td(cell_data_list[i][j], style="background-color:#FFDEAD;width:50px")
                elif i == 0 and j == 3:
                    mytr << td(cell_data_list[i][j], style="background-color:#FFDEAD;width:120px")
                elif i == 1 and j != 0:
                    mytr << td(cell_data_list[i][j], style="font-weight:bold")
                elif j == 0 and cell_data_list[i][j] == True:
                    mytr << td(cell_data_list[i][j], style="background-color:#FFB6C1;color:#800000")
                elif (j == 0 and cell_data_list[i][j] == False) or (i == 1 and j == 0):
                    mytr << td(cell_data_list[i][j])
                else:
                    mytr << td(cell_data_list[i][j], style="background-color:#FFDEAD")
        page << mytab

    if not os.path.exists(os.getcwd() + os.sep + 'html_result'):
        os.makedirs(os.getcwd() + os.sep + 'html_result')

    page.printOut(os.getcwd() + os.sep + 'html_result' + os.sep + sheet_name + '_html_data.html')


if __name__ == '__main__':
    start = time.time()
    # create_save_miss_html(sheet_name='ExistingSi', Silver_url_list=Silver_url_list, purl_bak_string=PURL_BAK_STRING)
    Silver_url_list = get_url_list_by_keyword(PURL_BAK_STRING, 'Silver')
    create_save_miss_html(sheet_name='Save-Miss', Silver_url_list=Silver_url_list, purl_bak_string=PURL_BAK_STRING)
    print time.time() - start

