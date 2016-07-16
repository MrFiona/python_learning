#!/usr/bin/env python 

import os 
import sys
import pexpect 

cmd = "scp 1.sh root@192.168.0.11:/home"
print cmd
child = pexpect.spawn(cmd)
child.expect("*(yes/no)?")
child.sendline("yes")
child.expect(".ssword:")
child.sendline("qq08061635")
child.interact()
