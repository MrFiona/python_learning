#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
        开发出一个日志系统，既要把日志输出到控制台，还要写入日志文件
"""

import os
import time
import os.path
import logging
from logging import NOTSET, DEBUG, INFO, WARNING, ERROR, CRITICAL
from machine_config import MachineConfig
CONFIG_FILE_PATH = os.getcwd() + os.sep + 'machineConfig' + os.sep + 'machine.conf'

class WorkLogger():
    def __init__(self, log_filename=None, log_level=logging.INFO):
        self.log_filename = log_filename
        self.log_level = log_level

        '''
                指定保存日志的文件路径，日志级别以及调用文件
                将日志存入到指定的文件中
        '''
    def print_message(self, msg=''):
        #读取配置文件日志、时间设置
        conf = MachineConfig(CONFIG_FILE_PATH)
        time_format = conf.get_node_info('logFormat', 'time_format')
        log_format = conf.get_node_info('logFormat', 'log_format')

        # 设置日志文件名称：time.time()取得当前时间；time.localtime()取得本地时间；time.strftime()格式化日期
        time_str = time.strftime('%s' %(time_format), time.localtime(time.time()))
        log_name = time_str + '_' + self.log_filename

        # 设置日志文件所在的路径
        log_file_dir = 'Log'
        if not os.path.isdir(log_file_dir):
            print("日志文件夹 %s 不存在，开始创建此文件夹" % log_file_dir)
            os.mkdir('Log')
        else:
            print("日志文件夹 %s 存在" % log_file_dir)

        os.chdir('Log')

        # 创建一个logger以及设置日志级别
        self.logger = logging.getLogger(self.log_filename)
        self.logger.setLevel(logging.DEBUG)

        # 创建文件handler，用于写入日志文件并设置文件日志级别
        file_handler = logging.FileHandler(log_name)
        file_handler.setLevel(logging.DEBUG)

        # 创建控制端输出handler，用于输出到控制端并设置输出日志级别
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)

        # 定义handler的输出格式并将格式应用到handler
        formatter = logging.Formatter(log_format)
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # 将handler加入到logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

        #根据日志级别打印信息
        if self.log_level == DEBUG or self.log_level == NOTSET:
            self.logger.debug(msg)
        elif self.log_level == INFO:
            self.logger.info(msg)
        elif self.log_level == WARNING:
            self.logger.warning(msg)
        elif self.log_level == ERROR:
            self.logger.error(msg)
        elif self.log_level == CRITICAL:
            self.logger.critical(msg)
        else:
            raise ValueError("日志级别参数有误[ %d ]" %(self.log_level))

        # 将handler从logger中移除
        self.logger.removeHandler(file_handler)
        self.logger.removeHandler(console_handler)

if __name__ == '__main__':
    _logger =  WorkLogger('Test Log', DEBUG)
    _logger.print_message('Hello！This is a test message！')