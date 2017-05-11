#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-25 16:07
# Author  : MrFiona
# File    : install_module.py
# Software: PyCharm Community Edition

import os

if __name__ == '__main__':
    print '\033[43m>>>>>>>>>> Start Installing Python Extra Modules <<<<<<<<<<\033[0m'
    'requirements.txt' in os.listdir(os.getcwd()) and os.system('pip --proxy child-prc.intel.com:913 install -r requirements.txt')
    print '\033[43m>>>>>>>>>>>>>>>>>>>>>>>    END    <<<<<<<<<<<<<<<<<<<<<<<<<\033[0m'