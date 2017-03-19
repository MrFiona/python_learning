#!/usr/bin/env python
# -*- coding: utf-8 -*-
# User    : apple
# File    : decorator_wrap.py
# Author  : MrFiona 一枚程序员
# Time     : 2017-03-19 22:44


import functools

def log(obj):
    if isinstance(obj,str):
        text=obj
        def decorator(func):
            @functools.wraps(func)
            def wrapper(*args, **kw):
                print 'begin %s %s():' % (text,func.__name__)
                func(*args, **kw)
                print 'args: ', args
                print 'kwargs: ', kw
                print 'end %s %s():' % (text,func.__name__)
            return wrapper
        return decorator

    else:
        func=obj
        print func
        @functools.wraps(func)
        def wrapper(*args, **kw):
            print 'begin111 %s():' % (func.__name__)
            func(*args, **kw)
            print 'args: ', args
            print 'kwargs: ', kw
            print 'end111 %s():' % (func.__name__)
        return wrapper
# @log('hello')
# def my_test(ID,student='MrFiona'):
#     print 'my_test run'

@log
def my_test1(ID,student='MrFiona'):
    print 'my_test run'

# my_test(ID=1024, student='John')
my_test1(ID=1024, student='John')


def decorator(func):
    @functools.wraps(func)
    def wrapper(*args, **kw):
        print 'begin %s():' % (func.__name__)
        func(*args, **kw)
        print 'args: ', args
        print 'kwargs: ', kw
        print 'end %s():' % (func.__name__)

    return wrapper

@decorator
def ggg():
    print 'ggggg'

ggg()