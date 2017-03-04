#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
        机器学习模型的操作配置文件模块
"""

import os
from exceptions import AttributeError, UserWarning
from ConfigParser import NoSectionError, RawConfigParser

#设置配置文件路径
CONFIG_FILE_PATH = os.getcwd() + os.sep + 'machineConfig' + os.sep + 'machine.conf'

class Machineconfig():
    def __init__(self, config_file_path=CONFIG_FILE_PATH):
        self.config_file_path = config_file_path

    def get_node_info(self, master_node=None, child_node=None):
        #节点信息和配置文件路径不能为空
        if not master_node or not child_node:
            raise UserWarning('请填写完整节点参数信息!!!')
        if not self.config_file_path:
            raise UserWarning('请指定配置文件路径!!!')

        conf = RawConfigParser()
        conf_status = conf.read(self.config_file_path)
        #验证配置文件路径是否正确
        if conf_status:
            try:
                child_node_format = conf.get(master_node, child_node)

            except NoSectionError as e:
                raise UserWarning('Please check the machine.conf file,section[ %s ] '
                                  'is not exist!' % (e.section))
            except AttributeError:
                raise UserWarning('配置文件节点信息未找到!!!')
        else:
            raise UserWarning('请检查配置文件路径是否正确!!!')

        return child_node_format

if __name__ == '__main__':
    m = Machineconfig()
    info = m.get_node_info('logFormat','time_format')
    print info
