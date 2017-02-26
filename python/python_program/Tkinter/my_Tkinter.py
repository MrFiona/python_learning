#!/usr/bin/env python 
#-*-:coding:utf-8-*-

'''实现Tkinter模块的相关功能
'''
from Tkinter import *
import os

def func():
	print "gggg!"

master = Tk()
#photo = PhotoImage(file=os.getcwd() + "/" + "11.png")
#v = StringVar()
#label = Label(master, text="hello! my name is zhangpeng!",image=photo)
#label = Label(master, text="hello! my name is zhangpeng!",textvariable=v)
#label = Label(master, text="hello! my name is zhangpeng!\noh! my god!",width=10,height=20,\
#	bg="white",fg="black",justify=LEFT,padx=200,command=func)

f = Frame(master)
#f.pack_propagate(0)
f.pack()
button = Button(f, text="hello!my name is zhangpeng!",command=func)
#v.set("新的文本")
#v.set("ahhhaa")
#label.pack()
#label.mainloop()
button.pack(fill=BOTH,expand=1)
#button.config(relief=SUNKEN)
button.mainloop()

