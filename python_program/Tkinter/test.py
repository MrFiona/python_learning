#!/usr/bin/env python 

import socket
import subprocess

def getHostName(ip):
    command = 'java -jar %s %s "hostname > %s.hostname"' %(remoteCmdLoca,ip,ip)
    result = subprocess.call(command, shell=True)
    command = '%s -q -r -pw passwd %s root@%s:/root' % (pscpLoca, pscpLoca, ip)
    result = subprocess.call(command, shell=True)
    command = '%s -q -r -pw passwd root@%s:/root/%s.hostname %s' %(pscpLoca,ip,ip,fileDir)
    result = subprocess.call(command, shell=True)
    fileName = fileDir + ip + '.hostname'
    readFile = open(fileName,'r')
    hostnameInfo =  str(readFile.readline().strip('\n'))
    readFile.close()
    subprocess.call('rm '+ fileName, shell=True)
    print "=========%s hostname is %s========" %(ip,hostnameInfo)
    return hostnameInfo

#getHostName("192.168.0.13")

print socket.gethostbyname("CentOS")
#print socket.gethostbyaddr("192.168.0.13")
print socket.getaddrinfo("CentOS","192.168.0.13")

