#!/usr/bin/env python
# -*- coding: utf-8 -*- #注意：此行代码是支持中文注释的条件

print "test 1"
def function(a, b):
    print a, b

apply(function,(1, 2+3))
apply(function,("hello world!", "my name is fiona!"))

plus_a=10
plus_b=2.0
print "result=",plus_a/plus_b,"gg"

my_name="Fionas,dsad"

print "oh,let's talk about %r qweqwe" %my_name 

formatter="%s %s %s %s"
print formatter % (1, 2, 3, 4)
print formatter % ("one", "two", "three", "four")


def printMax(x, y):
	'''print the maximum of two numbers.
	the two values must be intergers.'''

	x = int(x)
	y = int(y)
	
	if x > y:
		print x, 'is max'
	else:
		print y, 'is max'

printMax(3, 20)

print printMax.__doc__ 
#注意：doc两边分别有两个_线


  
  
