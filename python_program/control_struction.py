#!/usr/bin/python
# -*- coding: utf-8 -*- #注意：此行代码是支持中文注释的条件
#Filename:if.py

#number = "123"
number = 123
Running = True
#注意：此处的number为字符串123

while Running:
	#guess = raw_input('please input the num which you guess:')
	guess = input('please input the num which you guess:')
	#input返回的是数字

	#guess = int(raw_input('please input the num which you guess:'))
	#raw_input返回的是字符串,若number为数字,则需要int转换

	if guess == number:
		print 'Congratulations! you get it!'
		print "but you don't win any prize!"
		Running = False

	elif guess < number:
		print 'NO! it is a little lower than that!',
		#print后跟,号表示不换行

	else:
		print 'No! it is a little higher than that!',

else:
	print 'the while loop is over!'
#while有一个可选的else从句

print 'Done!'


#注意:python的缩进
