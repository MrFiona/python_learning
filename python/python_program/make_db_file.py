#!/usr/bin/env python 
# -*- coding:utf-8 -*-

from initdata import db

dbfilename = 'people_file.txt'
#将数据保存到文件中
def storeDbase(db, db_name=dbfilename):
	dbfile = open(dbfilename, 'w')
	for key in db:
		print(key,file=dbfile)
		for (name, value) in db[key].items():
			print(name, "=>", repr(value), file=dbfile)
		print("endrec", file=dbfile)
	print("enddb", file=dbfile)
	dbfile.close()

def loadDbase(db_name=dbfilename):
#读取文件内容到db字典中
	dbfile = open(dbfilename)
	import sys
	#sys.stdin = dbfile
	db = {}
	#key = input()
	key = dbfile.readline().rstrip('\n')
	while key != 'enddb': #名字
		rec = {}
		#file = input()	
		file = dbfile.readline().rstrip('\n')
		while file != 'endrec':
			name, value = file.split('=>')
			name = name.strip(' ')
			value = value.strip(' ')
			print("value:", value, type(value), type(eval(value)))
			rec[name] = eval(value)
			#print("value:", value, type(value))
			#file = input()	
			file = dbfile.readline().rstrip('\n')
		db[key] = rec
		#key = input()
		key = dbfile.readline().rstrip('\n')
	dbfile.close()
	return db

#将数据保存到文件中
storeDbase(db)

#读取文件内容到db字典中
db = loadDbase(dbfilename)
db['sue']['pay'] *= 1.10
db['tom']['name'] = 'Mary'

#将修改后数据保存到文件中
storeDbase(db)

for key in db:
	print(key, '=>\n', db[key])

file_fd = open('www.txt', 'w')
file_fd.write("hello this is test txt!")
file_fd.close()

file_fd = open(dbfilename, 'a+')
file_fd.write("hello this is test txt!")
file_fd.close()

file_fd = open(dbfilename, 'r')
cmd = file_fd.readline().rstrip('\n')
while len(cmd) != 0:
	#print(cmd,end='')
	print(cmd)
	cmd = file_fd.readline().rstrip('\n')
file_fd.close()

#if __name__ == '__main__':
#	from initdata import db
#	storeDbase(db)

