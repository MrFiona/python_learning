#!/usr/bin/env python 
# -*- coding:utf-8 -*-

import os
import time
#备份的文件默认放置的路径
DEFAULTS_DIR = "/mnt/hgfs/my_share_file"
#定义备份时间函数
def  str_time():
	return time.strftime("%Y_%m_%d_%H_%M_%S", time.localtime(time.time()))

def FileBackup(source_dirname1, destination_dirname=DEFAULTS_DIR):
	#检查备份的目录是否存在
	if os.path.exists(source_dirname1) is True:
		print "The backup directory %s exist!" %source_dirname1
	else:
		print "The backup directory %s doesn't exist!" %source_dirname1
	file_list = os.listdir(source_dirname1)
	#os.chdir(source_dirname1)
	#os.chdir("..")
	#current_path = os.getcwd()
	print "The backup directory contains of the following file:"
	i = 1
	#列出备份目录下的文件
	for filename in file_list:
		print "[ %d ]:\t%s" %(i, filename)
		i += 1
	#target =  destination_dirname + os.sep + str_time() + '.zip'
	#备份的压缩包名
	target =  destination_dirname + os.sep + source_dirname1 + "_" + str_time() + '.tar.gz'
	#command = "zip -qr %s %s" %(target, source_dirname1)
	#备份的目录前面不能加路径，不然就直接备份名字是最高级的目录名,如：/mnt/python，则备份名为mnt了，里面含有python
	command = "tar -cvf %s %s" %(target, os.path.basename(source_dirname1))
	if os.system(command) == 0:
		print "Successful backup to", target

source_dirname= raw_input("PLease input the backup directory:")
#FileBackup("/mnt/hgfs/my_share_file/python_program")
FileBackup(source_dirname)
