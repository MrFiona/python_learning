#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
        开发出一个日志系统，既要把日志输出到控制台，还要写入日志文件
"""

import logging
import time
import os
import os.path
import ConfigParser
import traceback
from exceptions import UserWarning

CONFIG_FILE_PATH = os.getcwd() + os.sep + 'machineConfig' + os.sep + 'machine.conf'

class Logger():
    def __init__(self, log_filename=None):
        self.log_filename = log_filename

        '''
                指定保存日志的文件路径，日志级别以及调用文件
                将日志存入到指定的文件中
        '''
    def PrintMessage(self, msg=''):
        # 设置日志文件名称：time.time()取得当前时间；time.localtime()取得本地时间；time.strftime()格式化日期；
        time_format = None
        log_format = None
        #读取配置文件日志时间设置
        try:
            time_conf = ConfigParser.RawConfigParser()
            time_conf.read(CONFIG_FILE_PATH)
            time_format = time_conf.get('logFormat', 'time_format')
            log_format = time_conf.get('logFormat1', 'log_format')
            # print time_format
        except:
            raise UserWarning("Get configurtion's message failed！")
            # print traceback._some_str('hello!')

        time_str = time.strftime('%s' %(time_format), time.localtime(time.time()))
        log_name = time_str + '_' + self.log_filename + '.log'

        # 设置日志文件所在的路径
        log_filedir = 'Log'
        if not os.path.isdir(log_filedir):
            print("日志文件夹 %s 不存在，开始创建此文件夹" % log_filedir)
            os.mkdir('Log')
        else:
            print("日志文件夹 %s 存在" % log_filedir)

        os.chdir('Log')

        # 创建一个logger以及设置日志级别
        # logging有6个日志级别:NOTSET, DEBUG, INFO, WARNING, ERROR, CRITICAL对应的值分别为：0,10,20,30,40,50
        # 例如:logging.DEBUG和10是等价的表示方法
        # 可以给日志对象(Logger Instance)设置日志级别，低于该级别的日志消息将会被忽略，也可以给Hanlder设置日志级别
        # 对于低于该级别的日志消息, Handler也会忽略。
        self.logger = logging.getLogger(self.log_filename)
        self.logger.setLevel(logging.DEBUG)

        # 创建文件handler，用于写入日志文件并设置文件日志级别
        file_handler = logging.FileHandler(log_name)
        file_handler.setLevel(logging.DEBUG)

        # 创建控制端输出handler，用于输出到控制端兵设置输出日志级别
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)

        # 定义handler的输出格式并将格式应用到handler
        formatter = logging.Formatter(log_format)
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # 将handler加入到logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

        # self.logger.debug(msg)
        self.logger.info(msg)

        # 将handler从logger中移除
        self.logger.removeHandler(file_handler)
        self.logger.removeHandler(console_handler)

if __name__ == '__main__':
    _logger = Logger('Test Log')
    _logger.PrintMessage('Hello！This is a test message！')


























































import matplotlib.pyplot as plt

# listX = [] #保存X轴数据
# listY = [] #保存Y轴数据
# listY1 = [] #保存Y轴数据
#
# file = file(r"C:\Users\apple\PycharmProjects\machine_program\TestFile\testlog.py","r")#打开日志文件
#
# while True:
#     line = file.readline()#读取一行日志
#     if len(line) == 0:#如果到达日志末尾，退出
#         break
#     paraList = line.split()
#     print paraList[2]
#     print paraList[3]
#     print paraList[4]
#     print paraList[5]
#     if paraList[2] == "Client0:": #在坐标图中添加两个点，它们的X轴数值是相同的
#         listX.append(float(paraList[3]))
#         listY.append(float(paraList[5]) - float(paraList[3]))
#         listY1.append(float(paraList[4]) - float(paraList[3]))
# file.close()


# plt.plot(listX,listY,'bo-',listX,listY1,'ro')#画图
#
# plt.title('tile')#设置所绘图像的标题
#
# plt.xlabel('time in sec')#设置x轴名称
#
# plt.ylabel('delays in ms')#设置y轴名称
# plt.show()