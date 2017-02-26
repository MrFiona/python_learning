#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
    1、从文本中提取数据
    2、将数据存放在新建文本
    3、将提取的数据用html表格显示
"""

#descriptions:
    # 1、脚本以及相关文件和原始数据必须放在同一个路径下
    # 2、提取出来的目标数据文件放在同级目录的dest_file目录下，名字定义为："expect_originalfilename"
    # 3、生成的html名字为week_text.html
    # 4、后续增加GUI界面实现参数控制的功能

import os
import sys
import re
import os.path
from pyh import *

reload(sys)
sys.setdefaultencoding('utf-8')

#define global variable
ORIGINAL_FILE_PATH = os.getcwd()
DESTINATION_FILE_PATH = os.getcwd() + os.sep + 'dest_file'

#从原始文本提提取数据
def fetchDataFromOrig(filename, pre_regex_word, back_regex_word, file_path=ORIGINAL_FILE_PATH, dest_path=DESTINATION_FILE_PATH):
    #判断目录的有效性

    if not os.path.exists(dest_path):
        os.mkdir(dest_path)

    if not os.path.isdir(file_path):
        raise ValueError('File Path Illegal!!!')
    #读取原始数据
    read_file_path = file_path + os.sep + '%s' %(filename)
    with open(read_file_path, 'r') as f:
        src_data_list = f.readlines()
    #正则提取数据 匹配多行
    # regex = re.compile(r'%s([\s\S]*?)%s' %(pre_regex_word, back_regex_word))
    regex = re.compile(r'%s(.*?)%s' %(pre_regex_word, back_regex_word), re.S)

    src_data_string = ''.join(src_data_list)
    object_data_list = regex.findall(src_data_string)
    if filename == 'OsInfo.txt':
        regex_1 = re.compile(r'\"WQL\"..(.*?)$', re.S)
        object_data_list_1 = regex_1.findall(src_data_string)
        object_data_list_1.insert(0, '"WQL", ')
        object_data_list.extend(object_data_list_1)
    print object_data_list
    object_data_string = ''.join(object_data_list)
    print object_data_string
    #将提取的数据存入文本
    with open(dest_path + os.sep + 'expect_%s' %(filename), 'w') as f:
        if filename == 'ApachePass.txt':
            f.write('Apache Pass' + object_data_string)
        elif filename == 'OsInfo.txt':
            f.write('OS Details' + object_data_string)
        elif filename == 'ProcessorInfo.txt':
            f.write('Genuine Intel(R) CPU 0000%@' + object_data_string)
        elif filename == 'SMBIOS.txt':
            f.write('BIOS Information {Type 0}' + object_data_string)
        else:
            raise ValueError('Save the data to the file faile!!!')

#从原始文本提提取数据
def fetchDataMain(file_path=ORIGINAL_FILE_PATH):
    #搜寻原始文件
    dir_file_list = os.listdir(file_path)
    for file in dir_file_list:
        if file == 'ApachePass.txt':
            fetchDataFromOrig('ApachePass.txt', 'Apache Pass', 'DIMMs')
        elif file == 'OsInfo.txt':
            fetchDataFromOrig('OsInfo.txt', 'OS Details', '"WQL"')
        elif file == 'ProcessorInfo.txt':
            fetchDataFromOrig('ProcessorInfo.txt', 'Genuine Intel\(R\) CPU 0000%@', 'Topology Details')
        elif file == 'SMBIOS.txt':
            fetchDataFromOrig('SMBIOS.txt', 'BIOS Information {Type 0}', 'BIOS Characteristics')

#将目标数据显示在html
class objectHtmlDataList:

    def __init__(self, dest_path=DESTINATION_FILE_PATH):
        self.dest_path = dest_path

    def commonTableFunction(self, filename, attribute_value_dict, throw_list):
        #tr 元素定义表格行，th 元素定义表头，td 元素定义表格单元
        t = table(border="1", cl="table1", cellpadding="0", cellspacing="0")
        t << tr(td('%s' % (filename), bgColor='#FFCCFF') + td(' %s Attribute Data' % (filename), bgColor='#CCFFFF'))
        for attribute, value in attribute_value_dict.items():
            t << tr(td(attribute) + td(value))
        t << th() + th()
        throw_list.append(t)

        return throw_list

    #生成html并显示数据
    def generateTable(self):
        # 存放属性和属性值对的字典
        attribute_value_dict = {}
        attribute_value_dict_info = {}
        attribute_value_dict_process = {}
        attribute_value_dict_bios = {}
        throw = []
        throw_info = []
        throw_process = []
        throw_bios = []

        dir_file_list = os.listdir(self.dest_path)

        for file in range(len(dir_file_list)):
            file_path = self.dest_path + os.sep + dir_file_list[file]
            # 提取目标文本中的属性值
            with open(file_path, 'r') as f:
                lines = f.readlines()
                line_list = ''.join(lines).splitlines()
                for line in range(len(line_list)):
                    if ',' in line_list[line]:
                        if dir_file_list[file] == 'expect_ApachePass.txt':
                            attribute_value_list = line_list[line].split(',')
                            attribute_value_dict[attribute_value_list[0].replace("\"""","")] = attribute_value_list[-1].replace("\"""","")
                        elif dir_file_list[file] == 'expect_OsInfo.txt':
                            attribute_value_list_info = line_list[line].split(',')
                            attribute_value_dict_info[attribute_value_list_info[0].replace("\"""", "")] = attribute_value_list_info[-1].replace("\"""", "")
                        elif dir_file_list[file] == 'expect_ProcessorInfo.txt':
                            attribute_value_list_process = line_list[line].split(',')
                            attribute_value_dict_process[attribute_value_list_process[0].replace("\"""","")] = attribute_value_list_process[-1].replace("\"""","")
                        elif dir_file_list[file] == 'expect_SMBIOS.txt':
                            attribute_value_list_bios = line_list[line].split(',')
                            attribute_value_dict_bios[attribute_value_list_bios[0].replace("\"""","")] = attribute_value_list_bios[-1].replace("\"""","")

        for file in range(len(dir_file_list)):
            if dir_file_list[file] == 'expect_ApachePass.txt':
                throw = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict, throw)
            elif dir_file_list[file] == 'expect_OsInfo.txt':
                throw_info = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict_info, throw_info)
            elif dir_file_list[file] == 'expect_ProcessorInfo.txt':
                throw_process = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict_process, throw_process)
            elif dir_file_list[file] == 'expect_SMBIOS.txt':
                throw_bios = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict_bios, throw_bios)

        page = PyH('Intel Compute Data report')
        page.addCSS('common.css')
        page << h1('Intel Compute Data report', align='center')

        tab = table(cellpadding="0", cellspacing="0", cl="table0")
        print len(throw)
        for t in range(0, len(throw), 2):
            tab << tr(td(throw[t - 1], cl="table0_td"))
        for t in range(0, len(throw_info), 2):
            tab << tr(td(throw_info[t - 1], cl="table0_td"))
        for t in range(0, len(throw_process), 2):
            tab << tr(td(throw_process[t - 1], cl="table0_td"))
        for t in range(0, len(throw_bios), 2):
            tab << tr(td(throw_bios[t - 1], cl="table0_td"))

        page << tab
        page.printOut('week_text.html')

if __name__ == '__main__':
    # fetchDataMain()
    # object = objectHtmlDataList()
    # object.generateTable()
    from collections import Iterator, Iterable
    import fileinput
    import glob
    for eachline in fileinput.FileInput(r"C:\Users\pengzh5x\PycharmProjects\machines_learning\ProcessorInfo.txt", backup='.bak', inplace=1):
        each = eachline.replace('Intel', 'hello')
        # sys.stdout.write('=>')
        sys.stdout.write(each)
        # # print fileinput.isfirstline()
        # print eachline
        # begainline = fileinput.lineno()
        # print begainline
        # print fileinput.lineno()

    # print isinstance(file, Iterable)
    # while file:
        # file.readline(1)
        # file = file.next()

    for line in fileinput.input(glob.glob(r'*.html')):
        if fileinput.isfirstline():
            print '-' * 20, 'Reading %s...' % fileinput.filename(), '-' * 20
        print str(fileinput.lineno()) + ': ' + line.upper(),

    print os.getcwd()