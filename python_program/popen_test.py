#!/usr/bin/env python
#-*- coding:utf-8 -*-
import os
def print_environment():
	i = 1
	print("program is start runing!\n")
	path_name = os.popen("ls ")#.readline()#.strip()
		
	while True:
		path_tmp = path_name.readline()
		print(path_tmp), #本身是有换行符的，这里要去掉print本身产生的换行符
		if len(path_tmp) == 0:
			break
		#for print_name in path_tmp:
		#	print("result %s:%s" %(i,print_name))
		#	i += 1
print_environment()

import threading, time, random
count = 0
class Counter(threading.Thread):
    def __init__(self, lock, threadName):
        '''@summary: 初始化对象。
        
        @param lock: 琐对象。
        @param threadName: 线程名称。
        '''
        super(Counter, self).__init__(name = threadName)  #注意：一定要显式的调用父类的初始化函数。
        self.lock = lock
    
    def run(self):
        '''@summary: 重写父类run方法，在线程启动后执行该方法内的代码。
        '''
        global count
        self.lock.acquire()
        for i in xrange(10000):
            count = count + 1
        self.lock.release()
lock = threading.Lock()
for i in range(5): 
    Counter(lock, "thread-" + str(i)).start()
time.sleep(2)	#确保线程都执行完毕
print count

if __name__ == '__main__':
	print("this is popen_test.py programe!\n")

