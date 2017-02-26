#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import ConfigParser
# print __file__

print os.getcwd()
os.chdir('machineConfig')

#遇到特殊字符的读写用RawConfigParser()
conf = ConfigParser.RawConfigParser()
conf.read('machine.conf')
sec = conf.sections()

print '\033[43m sections_list：%s\033[0m' %(conf.sections())
for opt in sec:
    print type(opt)
    print conf.items(opt)
    print '%s  %s' %(opt, conf.options(opt))

print conf.get('logFormat', 'time_format')

