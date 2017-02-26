#!/usr/bin/env python
# -*- coding: utf-8 -*-

# """
#     1、从文本中提取数据
#     2、将数据存放在新建文本
#     3、将提取的数据用html表格显示
# """
#
# import os
# import re
# import os.path
# from pyh import *
# import collections
#
# ORIGINAL_FILE_PATH = 'C:\Users\pengzh5x\Desktop\windows'
# DESTINATION_FILE_PATH = 'C:\Users\pengzh5x\Desktop\windows\dest_file'
#
# #从原始文本提提取数据
# def fetchDataFromOrig(filename, pre_regex_word, back_regex_word, file_path=ORIGINAL_FILE_PATH, dest_path=DESTINATION_FILE_PATH):
#     #判断目录的有效性
#     if not os.path.isdir(file_path):
#         raise ValueError('File Path Illegal!!!')
#     # if not os.path.exists(dest_path):
#     #     os.mkdir(dest_path)
#     #     os.chmod(dest_path, 0777)
#     #读取原始数据
#     read_file_path = file_path + os.sep + '%s' %(filename)
#     with open(read_file_path, 'r') as f:
#         src_data_list = f.readlines()
#     # print src_data_list
#     #正则提取数据
#     regex = re.compile(r'%s([\s\S]*?)%s' %(pre_regex_word, back_regex_word))
#     src_data_string = ''.join(src_data_list)
#     # print src_data_string
#     object_data_list = regex.findall(src_data_string)
#     # print object_data_list
#     object_data_string = ''.join(object_data_list)
#     # print object_data_string
#     #将提取的数据存入文本
#     with open(dest_path + os.sep + 'expect_%s' %(filename), 'w') as f:
#         if filename == 'ApachePass.txt':
#             f.write('Apache Pass' + object_data_string)
#         elif filename == 'OsInfo.txt':
#             f.write('OS Details' + object_data_string)
#         elif filename == 'ProcessorInfo.txt':
#             f.write('Genuine Intel(R) CPU 0000%@' + object_data_string)
#         elif filename == 'SMBIOS.txt':
#             f.write('BIOS Information {Type 0}' + object_data_string)
#         else:
#             raise ValueError('Save the data to the file faile!!!')
#
# #从原始文本提提取数据
# def fetchDataMain(file_path=ORIGINAL_FILE_PATH):
#     #搜寻原始文件
#     dir_file_list = os.listdir(file_path)
#     print dir_file_list
#     for file in dir_file_list:
#         if file == 'ApachePass.txt':
#             fetchDataFromOrig('ApachePass.txt', 'Apache Pass', 'DIMMs')
#         elif file == 'OsInfo.txt':
#             fetchDataFromOrig('OsInfo.txt', 'OS Details', '\n')
#         elif file == 'ProcessorInfo.txt':
#             fetchDataFromOrig('ProcessorInfo.txt', 'Genuine Intel\(R\) CPU 0000%@', 'Topology Details')
#         elif file == 'SMBIOS.txt':
#             fetchDataFromOrig('SMBIOS.txt', 'BIOS Information {Type 0}', 'BIOS Characteristics')
#
# #将目标数据显示在html
# class objectHtmlDataList:
#
#     def __init__(self, pro='Data List Using Html', dest_path=DESTINATION_FILE_PATH):
#         self.pro = pro
#         self.dest_path = dest_path
#
#     def lineWordsDeal(self, filename, line, line_list):
#         # if attribute_value_dict == None:
#         attribute_value_dict = {}
#         if ',' in line_list[line]:
#             attribute_value_list = line_list[line].split(',')
#             # 去除双引号
#             attribute_value_dict[attribute_value_list[0].replace("\"""", "")] = attribute_value_list[-1].replace("\"""", "")
#             # print 'sss %s' %(attribute_value_dict)
#             return attribute_value_dict
#
#     def commonTableFunction(self, filename, attribute_value_dict, throw_list):
#         print attribute_value_dict
#         t = table(caption='%s' % (filename), border="1", cl="table1", cellpadding="0", cellspacing="0")
#         t << tr(td('%s' % (filename), bgColor='#0099ff') + td(' %s Attribute Data' % (filename), bgColor='#0023ff'))
#         for attribute, value in attribute_value_dict.items():
#             t << tr(td(attribute) + td(value))
#         throw_list.append(t)
#
#         return throw_list
#
#     #生成html并显示数据
#     def generateTable(self):
#         # 存放属性和属性值对的字典
#         attribute_value_dict = {}
#         attribute_value_dict_info = {}
#         attribute_value_dict_process = {}
#         attribute_value_dict_bios = {}
#         throw = []
#         throw_info = []
#         throw_process = []
#         throw_bios = []
#         dir_file_list = os.listdir(self.dest_path)
#
#         for file in range(len(dir_file_list)):
#             file_path = self.dest_path + os.sep + dir_file_list[file]
#             # 提取目标文本中的属性值
#             with open(file_path, 'r') as f:
#                 lines = f.readlines()
#                 line_list = ''.join(lines).splitlines()
#                 for line in range(len(line_list)):
#                     if ',' in line_list[line]:
#                         if dir_file_list[file] == 'expect_ApachePass.txt':
#                             # attribute_value_dict = \
#                             #     self.lineWordsDeal(dir_file_list[file], line, line_list)
#                             # print 'ww %s' %(attribute_value_dict)
#                             attribute_value_list = line_list[line].split(',')
#                             # 去除双引号
#                             attribute_value_dict[attribute_value_list[0].replace("\"""","")] = attribute_value_list[-1].replace("\"""","")
#                         elif dir_file_list[file] == 'expect_OsInfo.txt':
#                             # attribute_value_dict_info = \
#                             #     self.lineWordsDeal(dir_file_list[file], line, line_list)
#                             attribute_value_list_info = line_list[line].split(',')
#                             # 去除双引号
#                             attribute_value_dict_info[attribute_value_list_info[0].replace("\"""", "")] = attribute_value_list_info[-1].replace("\"""", "")
#                         elif dir_file_list[file] == 'expect_ProcessorInfo.txt':
#                             # attribute_value_dict_process = \
#                             #     self.lineWordsDeal(dir_file_list[file], line, line_list)
#                             attribute_value_list_process = line_list[line].split(',')
#                             #去除双引号
#                             attribute_value_dict_process[attribute_value_list_process[0].replace("\"""","")] = attribute_value_list_process[-1].replace("\"""","")
#                         elif dir_file_list[file] == 'expect_SMBIOS.txt':
#                             # attribute_value_dict_bios = \
#                             #     self.lineWordsDeal(dir_file_list[file], line, line_list)
#                             attribute_value_list_bios = line_list[line].split(',')
#                             #去除双引号
#                             attribute_value_dict_bios[attribute_value_list_bios[0].replace("\"""","")] = attribute_value_list_bios[-1].replace("\"""","")
#
#         for file in range(len(dir_file_list)):
#             if dir_file_list[file] == 'expect_ApachePass.txt':
#                 throw = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict, throw)
#             elif dir_file_list[file] == 'expect_OsInfo.txt':
#                 throw_info = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict_info, throw_info)
#             elif dir_file_list[file] == 'expect_ProcessorInfo.txt':
#                 throw_process = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict_process, throw_process)
#             elif dir_file_list[file] == 'expect_SMBIOS.txt':
#                 throw_bios = self.commonTableFunction(dir_file_list[file].split('.')[0], attribute_value_dict_bios, throw_bios)
#
#         page = PyH('Intel Compute Data report')
#         page.addCSS('common.css')
#         page << h1('Intel Compute Data report', align='center')
#
#         tab = table(cellpadding="0", cellspacing="0", cl="table0")
#         for t in range(0, len(throw), 2):
#             tab << tr(td(throw[t - 1], cl="table0_td"))
#         for t in range(0, len(throw_info), 2):
#             tab << tr(td(throw_info[t - 1], cl="table0_td"))
#         for t in range(0, len(throw_process), 2):
#             tab << tr(td(throw_process[t - 1], cl="table0_td"))
#         for t in range(0, len(throw_bios), 2):
#             tab << tr(td(throw_bios[t - 1], cl="table0_td"))
#         page << tab
#         page.printOut('week_text.html')
#
# fetchDataMain()
# object = objectHtmlDataList()
# object.generateTable()

#移除文件中的某一行数据并回写进文件
def removeLine(filename, lineno):
    fro = open(filename, "rb")

    current_line = 0
    while current_line < lineno:
        fro.readline()
        current_line += 1

    seekpoint = fro.tell()
    frw = open(filename, "r+b")
    frw.seek(seekpoint, 0)

    # read the line we want to discard
    fro.readline()

    # now move the rest of the lines in the file
    # one line back
    chars = fro.readline()
    while chars:
        frw.writelines(chars)
        chars = fro.readline()

    fro.close()
    frw.truncate()
    frw.close()
removeLine(r'C:\Users\pengzh5x\Desktop\windows\test.txt', 10)


#删除某一行数据
# data = open(r'C:\Users\pengzh5x\Desktop\windows\Apache_pass.txt', 'rt').readlines()
#
# with open(r'C:\Users\pengzh5x\Desktop\windows\Apache_pass.txt', 'wt') as handle:
#
#     handle.writelines(data[:6])
#
#     handle.writelines(data[6+1:])