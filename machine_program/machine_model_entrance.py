#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-12 17:28
# Author  : MrFiona
# File    : machine_model_entrance.py
# Software: PyCharm Community Edition

import os
import time
import traceback
from get_all_html import GetUrlFromHtml
from insert_excel import InsertDataIntoExcel
from cache_mechanism import DiskCache
from create_email_html import create_save_miss_html
from public_use_function import backup_excel_file, get_url_list_by_keyword
from send_email import SendEmail
from setting_gloab_variable import PURL_BAK_STRING, type_sheet_name_list

#模型执行入口函数
#默认是生成Purley-FPGA数据
def machine_model_entrance(purl_bak_string='Purley-FPGA'):
    # 获取html
    get_url_object = GetUrlFromHtml('https://dcg-oss.intel.com/ossreport/auto/')
    get_url_object.get_all_type_data()
    # 写文件
    get_url_object.write_html_by_multi_thread()

    cache = DiskCache()

    #将数据插入excel
    Silver_url_list = get_url_list_by_keyword(PURL_BAK_STRING, 'Silver')
    insert_object = InsertDataIntoExcel(verificate_flag=False, purl_bak_string=purl_bak_string, cache=cache, silver_url_list=Silver_url_list)
    func_name_list = insert_object.return_name().keys()
    call_func_list = [ func for func in func_name_list if func.startswith('insert') ]
    print call_func_list
    for func in call_func_list:
        getattr(insert_object, func)(purl_bak_string, 1)

    insert_object.close_workbook()

    # 生成html
    failed_sheet_name_list = []
    for type_name in type_sheet_name_list:
        try:
            create_save_miss_html(sheet_name=type_name, purl_bak_string=PURL_BAK_STRING,
                                  Silver_url_list=Silver_url_list)
        except:
            failed_sheet_name_list.append(type_name)
            print traceback.print_exc()

    print failed_sheet_name_list
    # 备份excel文件
    backup_excel_file()
    SendEmail()


if __name__ == '__main__':
    start_time = time.time()
    #执行之前安装所需模块
    os.system('python install_module.py')
    machine_model_entrance(PURL_BAK_STRING)
    end_time = time.time()
    print end_time - start_time