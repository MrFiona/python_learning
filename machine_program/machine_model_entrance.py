#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-12 17:28
# Author  : MrFiona
# File    : machine_model_entrance.py
# Software: PyCharm Community Edition

import os
import time
import shutil
import traceback
from multiprocessing.dummy import Pool as ThreadPool
from get_all_html import GetUrlFromHtml
from insert_excel import InsertDataIntoExcel
from cache_mechanism import DiskCache
from create_email_html import create_save_miss_html
from public_use_function import backup_excel_file
from send_email import SendEmail
from setting_gloab_variable import PURL_BAK_STRING, Silver_url_list, type_sheet_name_list

#模型执行入口函数
#默认是生成Purley-FPGA数据
def machine_model_entrance(purl_bak_string='Purley-FPGA', silver_url_list=None):
    cache = DiskCache()
    #获取html
    # get_url_object = GetUrlFromHtml('https://dcg-oss.intel.com/ossreport/auto/')
    # get_url_object.get_all_type_data()
    #写文件
    # get_url_object.write_html_by_multi_thread()

    #将数据插入excel
    insert_object = InsertDataIntoExcel(verificate_flag=False, purl_bak_string=purl_bak_string, cache=cache, silver_url_list=silver_url_list)
    func_name_list = insert_object.return_name().keys()
    call_func_list = [ func for func in func_name_list if func.startswith('insert') ]
    print call_func_list
    for func in call_func_list:
        getattr(insert_object, func)(purl_bak_string, 1)
        # pool = ThreadPool(11)
        # pool.map(getattr(insert_object, func), [10])
        # pool.close()
        # pool.join()

    insert_object.close_workbook()

    copy_path = os.getcwd() + os.sep + 'excel_dir'
    for file in type_sheet_name_list:
        shutil.copy2(copy_path + os.sep + PURL_BAK_STRING + '_report_result.xlsx',
                     copy_path + os.sep + PURL_BAK_STRING + '_report_result_' + file + '.xlsx')


if __name__ == '__main__':
    start_time = time.time()
    #执行之前安装所需模块
    os.system('python install_module.py')
    #备份excel文件
    machine_model_entrance(PURL_BAK_STRING, Silver_url_list)
    #生成html
    failed_sheet_name_list = []
    for type_name in type_sheet_name_list:
        try:
            create_save_miss_html(sheet_name=type_name, purl_bak_string=PURL_BAK_STRING, Silver_url_list=Silver_url_list)
        except:
            failed_sheet_name_list.append(type_name)
            print traceback.print_exc()

    print failed_sheet_name_list
    backup_excel_file()
    SendEmail()
    end_time = time.time()
    print end_time - start_time



# Traceback (most recent call last):
#   File "C:/Users/pengzh5x/Desktop/python_scripts/machine_model_entrance.py", line 51, in <module>
#     create_save_miss_html(sheet_name=type_name, purl_bak_string=PURL_BAK_STRING, Silver_url_list=Silver_url_list)
#   File "C:\Users\pengzh5x\Desktop\python_scripts\create_email_html.py", line 40, in create_save_miss_html
#     cell_data_list = get_report_data(sheet_name, purl_bak_string=purl_bak_string)
#   File "C:\Users\pengzh5x\Desktop\python_scripts\public_use_function.py", line 406, in get_report_data
#     data = win_book.getCell(sheet=sheet_name, row=i, col=j)
#   File "C:\Users\pengzh5x\Desktop\python_scripts\public_use_function.py", line 343, in getCell
#     sht = self.xlBook.Worksheets(sheet)
#   File "C:\Python27\lib\site-packages\win32com\client\dynamic.py", line 192, in __call__
#     return self._get_good_object_(self._oleobj_.Invoke(*allArgs),self._olerepr_.defaultDispatchName,None)
# pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2147352565), None)