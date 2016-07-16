#!/usr/bin/python

a=15;exec r"a-=1;b=abs(a);_='-*?@.#`%$\\:[{\'\"'[b:];print _[::-1]+_[0]*2*(b+9)+_;"*29

x = 1
print id(x)

mm = "/mnt/sadqe/bb"

import os 

print os.path.basename(mm)
print os.path.dirname(mm)

for file in os.path.split(mm):
	print file

bb = os.path.split(mm)
print bb
