#!/usr/bin/python
#-*- coding:utf-8 -*-
#自动连接ssh的代码 

import sys, time, os
import pexpect
 
#try:
#    import pexpect
#except ImportError:
#    print "You must install pexpect module" 
#    sys.exit(1)

addr_map = {
    'v3' :('root@192.168.0.13', 'qq08061635'),
    'dev':('test016@192.168.1.136', 'test016'),
}
 
try:
    key = sys.argv[1]
    host = addr_map[key]
except:
    print ("argv error, use it link jssh v3, v3 must defined in addr_map")
    sys.exit(1)
 
server = pexpect.spawn('/usr/bin/ssh %s' % host[0])
server.expect('.*ssword:')
server.sendline(host[1])
server.interact()
