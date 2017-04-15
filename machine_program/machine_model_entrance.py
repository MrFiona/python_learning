#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-12 17:28
# Author  : MrFiona
# File    : machine_model_entrance.py
# Software: PyCharm Community Edition

import time
from multiprocessing.dummy import Pool as ThreadPool
from get_all_html import GetUrlFromHtml
from insert_excel import InsertDataIntoExcel
from cache_mechanism import DiskCache

#模型执行入口函数
def machine_model_entrance():
    cache = DiskCache()
    #获取html
    get_url_object = GetUrlFromHtml('https://dcg-oss.intel.com/ossreport/auto/')
    get_url_object.get_all_type_data()
    # 多线程写文件
    get_url_object.write_html_by_multi_thread()

    #将数据插入excel
    insert_object = InsertDataIntoExcel(verificate_flag=False, cache=cache)
    func_name_list = insert_object.return_name().keys()
    call_func_list = [ func for func in func_name_list if func.startswith('insert') ]
    print call_func_list
    for func in call_func_list:
        getattr(insert_object, func)(1)
        # pool = ThreadPool(11)
        # pool.map(getattr(insert_object, func), [10])
        # pool.close()
        # pool.join()
    insert_object.close_workbook()

if __name__ == '__main__':
    start_time = time.time()
    machine_model_entrance()
    end_time = time.time()
    print end_time - start_time