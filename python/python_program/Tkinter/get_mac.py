#!/usr/bin/env python 
#-*- coding:utf-8 -*-

#(2)按照操作系统平台来
def get_mac_address():
    '''
    @summary: return the MAC address of the computer
    '''
    import sys
    import os
    mac = None
    if sys.platform == "win32":
        for line in os.popen("ipconfig /all"):
            #print line
            print (line)
            if line.lstrip().startswith("物理地址"):
            #if line.lstrip().startswith("Physical Address"):
                mac = line.split(":")[1].strip().replace("-", ":")
                break
    else:
        for line in os.popen("/sbin/ifconfig"):
            if 'Ether' in line:
                mac = line.split()[4]
                break
    return mac
#print (get_mac_address())
#print get_mac_address()

#(1)通用方法,借助uuid模块
def get_mac_address():
	import uuid
	node = uuid.getnode()
	mac = uuid.UUID(int = node).hex[-12:]
	mac_str = mac[0:2] + ":" + mac[2:4] + ":" + mac[4:6] + ":" + mac[6:8] + ":" + mac[8:10] + ":" + mac[10:12]
	mac_str = mac_str.upper()
	return mac_str

print get_mac_address()

#需要提醒的是，方法一，也就是通过uuid来实现的方法适用性更强一些，因为通过解析popen的数据还受到系统语言环境的影响
#比如在中文环境 下"ipconfig/all"的结果就包含有中文，要查找mac地址的话就不能查找Physical Address;而应该是“物理地址


