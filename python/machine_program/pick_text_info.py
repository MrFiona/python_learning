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
import re
import shutil
import sys
import glob
import ConfigParser
import collections
from pyh import *

reload(sys)
sys.setdefaultencoding('utf-8')

color_dict = {'ApachePass':'44', 'OsInfo':'45', 'ProcessorInfo':'46', 'SMBIOS':'47'}

#define global variable
ORIGINAL_FILE_PATH = os.path.split(os.getcwd())[0] + os.sep + 'orignal_text'
DESTINATION_FILE_PATH = os.path.split(os.getcwd())[0] + os.sep + 'result_text'
CONFIG_FILE_PATH = 'C:\Users\pengzh5x\PycharmProjects\machines_learning\get_text_task\machine.conf'

def removeline(filename, signal_lineno, continue_lineno):
    #删除单个行
    def signal_remove(lineno):
        # 移除文件中的某一行数据并回写进文件
        fro = open(filename, "rb")

        current_line = 0
        while current_line < lineno - 1:
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

    #倒着删除可以避免行移动
    signal_lineno.sort(reverse=True)

    for lineno in signal_lineno:
        signal_remove(lineno)

    continue_lineno_list = []
    for k in range(len(continue_lineno)):
        if len(continue_lineno[k]) != 0:
            #continue_line配置参数必须成对出现
            continue_lineno[k].sort(reverse=True)
            if len(continue_lineno[k])%2 != 0:
                raise UserWarning('请检查配置参数!!!continue_line配置有误!!!')

            for line in range(continue_lineno[k][1], continue_lineno[k][0]+1):
                continue_lineno_list.append(line)

    continue_lineno_list.sort(reverse=True)

    for lineno in continue_lineno_list:
        if lineno != 0:
            signal_remove(lineno)

def removeSignalContinueData(filename, pre_regex_word, head_regex_word,
                             config_file_path=CONFIG_FILE_PATH,
                             orignal_file_path=ORIGINAL_FILE_PATH):
    pre_flag = False
    head_flag = False
    signal_list = []
    continue_container = []

    conf = ConfigParser.RawConfigParser()
    conf.read(config_file_path)
    signal_string = conf.get('signal_line', filename.split('.')[0] + '_' + 'signal')
    continue_string = conf.get('continue_line', filename.split('.')[0] + '_' + 'continue')

    # 判断提取标记是否被包含在内
    f = open(orignal_file_path + os.sep + filename, 'r')
    line = f.readline()

    pre_num = 1
    head_num = 1

    while len(line) != 0:
        if len(pre_regex_word) != 0 and line.strip('\n') == pre_regex_word:
            pre_flag = True
        if len(head_regex_word) !=0 and line.strip('\n') == head_regex_word:
            head_flag = True
        line = f.readline()
        if not pre_flag:
            pre_num += 1
        if not head_flag:
            head_num += 1

    print pre_num, head_num

    if len(signal_string) != 0:
        signal_list = signal_string.split(',')
        signal_list = [eval(signal_list[k]) for k in range(len(signal_list))]
        pre_signal_flag_list = [True if i == pre_num else False for i in signal_list]
        head_signal_flag_list = [True if i == head_num else False for i in signal_list]
        if True in pre_signal_flag_list or True in head_signal_flag_list:
            raise UserWarning('不能将提取标记包含在删除行之内!!!删除标记行为 [%d]!!!' % (pre_num))
        signal_list = [num for num in signal_list if num != 0]
        signal_list = list(set(signal_list))
        signal_list.sort()

    #连续行出现多个需要考虑这种情况
    if len(continue_string) != 0:
        continue_list = continue_string.split(',')
        continue_line_list = \
            [continue_list[i].split('-') for i in range(len(continue_list))]

        # 字符串转化为整型
        continue_line_list = \
            [eval(continue_line_list[m][n]) for m in range(len(continue_line_list))
             for n in range(len(continue_line_list[m]))]

        #存储多个连续行数组
        continue_line_list_list = \
            [ [continue_line_list[j] for j in range(2*i,2*i+2)] for i in range(len(continue_line_list)/2) ]
        #在不考虑交集的情况下，只需要按照数组的第一位倒序排序即可，避免行动态移动
        continue_line_list_list.sort(key=lambda x: x[0], reverse=True)

        #处理多行之间出现交集的情况
        temp = []
        def range_list(num1, num2):
            temp_list = []
            for i in range(num1, num2):
                temp_list.append(i)
            return temp_list
        #temp存放多行列表中存在交集的数据对应的下标
        def mix_index_list(continue_line_list_list):
            for i in range(len(continue_line_list_list)):
                i_list = range_list(continue_line_list_list[i][0],continue_line_list_list[i][1])
                for j in range(i+1, len(continue_line_list_list)):
                    j_list = range_list(continue_line_list_list[j][0], continue_line_list_list[j][1])
                    if list(set(i_list) & set(j_list)):
                        temp.append([i,j])
            return temp

        temp = mix_index_list(continue_line_list_list)
        print 'temp:', temp, len(temp)

        while len(temp) != 0:
            #合并存在交集的行数据，添加进原始多行列表（后增，否则待删除的数据下标会变化）
            for i in range(len(temp)):
                link = [ min(continue_line_list_list[temp[i][0]][0], continue_line_list_list[temp[i][1]][0]),
                         max(continue_line_list_list[temp[i][0]][1], continue_line_list_list[temp[i][1]][1])]
                continue_line_list_list.append(link)
            #存储合并后的待删除原始数据
            m = []
            for i in range(len(temp)):
                for j in range(2):
                    m.append(continue_line_list_list[temp[i][j]])
            #删除多行列表中合并后的原始数据
            for k in range(len(m)):
                try:
                    continue_line_list_list.remove(m[k])
                #可能会出现同一个行列表与不止一个行列表相交的情况
                except ValueError as e:
                    pass

            temp = []
            continue_line_list_list.sort()
            temp = mix_index_list(continue_line_list_list)

        #判断是否将标记行包含在内
        for continue_line_list in continue_line_list_list:
            pre_continue_flag_list = [
                True if pre_num >= continue_line_list[i * 2] and pre_num <= continue_line_list[2 * i + 1]
                else False for i in range(len(continue_line_list) / 2)]
            head_continue_flag_list = [
                True if head_num >= continue_line_list[i * 2] and head_num <= continue_line_list[2 * i + 1]
                else False for i in range(len(continue_line_list) / 2)]
            if True in pre_continue_flag_list and len(pre_regex_word) != 0:
                raise UserWarning('不能将提取标记包含在删除行之内!!!删除标记行为 [%d]!!!' % (pre_num))
            if True in head_continue_flag_list and len(head_regex_word) != 0:
                raise UserWarning('不能将提取标记包含在删除行之内!!!删除标记行为 [%d]!!!' % (head_num))
            continue_line_list = list(set(continue_line_list))

            line_num_dict = collections.OrderedDict()
            for i in range(len(continue_line_list)):
                line_num_dict[chr(ord('A') + i)] = 0

            continue_line_list.sort(reverse=False)
            if len(signal_string) != 0:
                for num in range(len(signal_list)):
                    for nu in range(0, len(continue_line_list)/2):
                        if signal_list[num] > continue_line_list[nu*2] and signal_list[num] <= continue_line_list[nu*2+1]:
                            for j in range(nu*2+1, len(continue_line_list)):
                                num_dict_1 = chr(ord('A')+j)
                                line_num_dict.update({num_dict_1: line_num_dict[num_dict_1]-1})
                        elif signal_list[num] <= continue_line_list[nu*2]:
                            for i in range(nu*2, len(continue_line_list)):
                                num_dict_1 = chr(ord('A')+i)
                                line_num_dict.update({num_dict_1: line_num_dict[num_dict_1]-1})

                for i in range(len(continue_line_list)):
                    if line_num_dict[chr(ord('A')+i)] < 0:
                        continue_line_list[i] += line_num_dict[chr(ord('A')+i)]
                    elif line_num_dict[chr(ord('A')+i)] > 0:
                        continue_line_list[i] -= line_num_dict[chr(ord('A')+i)]
                        continue_line_list[i] -= line_num_dict[chr(ord('A')+i)]

            continue_container.append(continue_line_list)
            print 'signal_list:', signal_list
            print 'continue_list:', continue_line_list
    continue_container.sort(key=lambda x: x[0], reverse=True)
    print 'container: ',continue_container
    removeline(ORIGINAL_FILE_PATH + os.sep + filename, signal_list, continue_container)

def removeTempFile(pre_path):
    for file in glob.glob(pre_path + os.sep + '*.tmp'):
        os.remove(file)

def copyTempFile(pre_file_name, pre_path):
    shutil.copy2(pre_path + pre_file_name + '.txt', pre_path + pre_file_name + '.tmp')

def fetchDataFromOrig(filename, pre_regex_word, back_regex_word, file_path=ORIGINAL_FILE_PATH,
                      dest_path=DESTINATION_FILE_PATH):

    if not os.path.exists(dest_path):
        os.mkdir(dest_path)

    if not os.path.isdir(file_path):
        raise ValueError('File Path Illegal!!!')
    #读取原始数据
    pre_file_name = filename.split('.')[0]
    read_file_path = file_path + os.sep + '%s' %(filename)
    with open(read_file_path, 'r') as f:
        src_data_list = f.readlines()
    #正则提取数据 匹配多行
    # regex = re.compile(r'%s([\s\S]*?)%s' %(pre_regex_word, back_regex_word))
    regex = re.compile(r'%s(.*)%s\n' %(pre_regex_word, back_regex_word), re.S)

    src_data_string = ''.join(src_data_list)
    object_data_list = regex.findall(src_data_string)
    # if filename == 'OsInfo.txt':
    #     regex_1 = re.compile(r'\"WQL\"..(.*?)$', re.S)
    #     object_data_list_1 = regex_1.findall(src_data_string)
    #     object_data_list_1.insert(0, '"WQL", ')
    #     object_data_list.extend(object_data_list_1)
    # print object_data_list
    object_data_string = ''.join(object_data_list)
    print '\033[%dm+**********************************************************+\033[0m' %(eval(color_dict[pre_file_name]))
    print '文本 [%s] 提取出的内容如下:' %(filename)
    print object_data_string.strip()
    print '\033[%dm+**********************************************************+\033[0m\n' %eval((color_dict[pre_file_name]))

    with open(dest_path + os.sep + 'expect_%s' %(pre_file_name + '.txt'), 'w') as f:
        if pre_file_name == 'ApachePass':
            f.write('Apache Pass' + object_data_string)
        elif pre_file_name == 'OsInfo':
            f.write('OS Details' + object_data_string)
        elif pre_file_name == 'ProcessorInfo':
            f.write('Genuine Intel(R) CPU 0000%@' + object_data_string)
        elif pre_file_name == 'SMBIOS':
            f.write('BIOS Information {Type 0}' + object_data_string)
        else:
            raise ValueError('Save the data to the file faile!!!')


def fetchDataMain(file_path=ORIGINAL_FILE_PATH, config_file_path=CONFIG_FILE_PATH,
                  orignal_file_path=ORIGINAL_FILE_PATH):

    #原始文本文件所在的目录
    dir_file_list = os.listdir(file_path)
    pre_path = file_path + os.sep

    for file in dir_file_list:
        if file == 'ApachePass.txt':
            #拷贝原始数据副本文件做相应操作，防止损坏原始文件
            copyTempFile('ApachePass', pre_path)
            removeSignalContinueData('ApachePass.tmp', 'Apache Pass', 'DIMMs',
                                     config_file_path, orignal_file_path)
            #从副本中提取目标数据
            fetchDataFromOrig('ApachePass.tmp', 'Apache Pass', 'DIMMs')

        elif file == 'OsInfo.txt':
            copyTempFile('OsInfo', pre_path)
            removeSignalContinueData('OsInfo.tmp', 'OS Details', '',
                                     config_file_path, orignal_file_path)
            fetchDataFromOrig('OsInfo.tmp', 'OS Details', '')

        elif file == 'ProcessorInfo.txt':
            copyTempFile('ProcessorInfo', pre_path)
            removeSignalContinueData('ProcessorInfo.tmp', 'Genuine Intel\(R\) CPU 0000%@', 'Topology Details',
                                     config_file_path, orignal_file_path)
            fetchDataFromOrig('ProcessorInfo.tmp', 'Genuine Intel\(R\) CPU 0000%@', 'Topology Details')

        elif file == 'SMBIOS.txt':
            copyTempFile('SMBIOS', pre_path)
            removeSignalContinueData('SMBIOS.tmp', 'BIOS Information {Type 0}', 'BIOS Characteristics',
                                     config_file_path, orignal_file_path)
            fetchDataFromOrig('SMBIOS.tmp', 'BIOS Information {Type 0}', 'BIOS Characteristics')

    #移除临时文件
    # removeTempFile(pre_path)

#将目标数据显示在html
class objectHtmlDataList:

    def __init__(self, dest_path=DESTINATION_FILE_PATH):
        self.dest_path = dest_path
        self.throw = []
        self.throw_info = []
        self.throw_process = []
        self.throw_bios = []
        self.attribute_value_dict = {}
        self.attribute_value_dict_info = {}
        self.attribute_value_dict_process = {}
        self.attribute_value_dict_bios = {}

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
                            self.attribute_value_dict[attribute_value_list[0].replace("\"""","")] = attribute_value_list[-1].replace("\"""","")
                        elif dir_file_list[file] == 'expect_OsInfo.txt':
                            attribute_value_list_info = line_list[line].split(',')
                            self.attribute_value_dict_info[attribute_value_list_info[0].replace("\"""", "")] = attribute_value_list_info[-1].replace("\"""", "")
                        elif dir_file_list[file] == 'expect_ProcessorInfo.txt':
                            attribute_value_list_process = line_list[line].split(',')
                            self.attribute_value_dict_process[attribute_value_list_process[0].replace("\"""","")] = attribute_value_list_process[-1].replace("\"""","")
                        elif dir_file_list[file] == 'expect_SMBIOS.txt':
                            attribute_value_list_bios = line_list[line].split(',')
                            self.attribute_value_dict_bios[attribute_value_list_bios[0].replace("\"""","")] = attribute_value_list_bios[-1].replace("\"""","")

        for file in range(len(dir_file_list)):
            if dir_file_list[file] == 'expect_ApachePass.txt':
                self.throw = self.commonTableFunction(dir_file_list[file].split('.')[0], self.attribute_value_dict, self.throw)
            elif dir_file_list[file] == 'expect_OsInfo.txt':
                self.throw_info = self.commonTableFunction(dir_file_list[file].split('.')[0], self.attribute_value_dict_info, self.throw_info)
            elif dir_file_list[file] == 'expect_ProcessorInfo.txt':
                self.throw_process = self.commonTableFunction(dir_file_list[file].split('.')[0], self.attribute_value_dict_process, self.throw_process)
            elif dir_file_list[file] == 'expect_SMBIOS.txt':
                self.throw_bios = self.commonTableFunction(dir_file_list[file].split('.')[0], self.attribute_value_dict_bios, self.throw_bios)

        page = PyH('Intel Compute Data report')
        page.addCSS('common.css')
        page << h1('Intel Compute Data report', align='center')

        tab = table(cellpadding="0", cellspacing="0", cl="table0")
        for t in range(0, len(self.throw), 2):
            tab << tr(td(self.throw[t - 1], cl="table0_td"))
        for t in range(0, len(self.throw_info), 2):
            tab << tr(td(self.throw_info[t - 1], cl="table0_td"))
        for t in range(0, len(self.throw_process), 2):
            tab << tr(td(self.throw_process[t - 1], cl="table0_td"))
        for t in range(0, len(self.throw_bios), 2):
            tab << tr(td(self.throw_bios[t - 1], cl="table0_td"))

        page << tab
        page.printOut('week_text.html')


def remove_func(line_list, **kwargs):
    print 'args: ', line_list
    for j in kwargs:
        print 'kwargs: %s  %s' %(j, kwargs[j])


if __name__ == '__main__':
    # fetchDataMain()
    object = objectHtmlDataList()
    # object.generateTable()
    print object.__class__
    print object