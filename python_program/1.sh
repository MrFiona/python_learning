#!/usr/bin/env expect  
  
set timeout 30  
  
spawn ssh -l username 192.168.1.1  
  
expect "password:"  
  
send "ispass\r"  
  
interact  
  
##############################################  
  
  
