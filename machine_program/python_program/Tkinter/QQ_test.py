#!/usr/bin/env python 
#coding:utf-8

import time,datetime     
import urllib2    
def chk_qq(qqnum):  
        chkurl = 'http://wpa.paipai.com/pa?p=1:'+`qqnum`+':17'  
        a = urllib2.urlopen(chkurl)     
        length=a.headers.get("content-length")     
        a.close()    
        if length=='2348':     
            return 'Online'    
        elif length=='2205':     
            return 'Offline'    
        else:     
            return 'Unknown Status!'   
def writestate(statenow):  
        f=open(str(qq),'a')#将监测数据记录于同一文件夹下以QQ号命名的文件里  
        m=str(datetime.datetime.now())+"====state===="+statenow+"\n\r"  
        print m  
        f.write(m)  
        f.close()  
       
qq = "1160177283"#在此更改QQ号  
      
if __name__=='__main__':  
        while 1:
		stat = chk_qq(qq)  
		writestate(stat)  
		time.sleep(300)#5分钟测一次  
