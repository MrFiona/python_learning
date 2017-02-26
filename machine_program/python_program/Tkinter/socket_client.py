#socket_client.py
#coding:utf-8 
import socket
 
def socket_send(command):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.connect(('192.168.0.13', 1000))
    sock.send(command)
    result = sock.recv(2048)
    sock.close()
    return result
 
if __name__ == '__main__':
     print socket_send('ls && rm zhang.txt')
#sock.connect(address)，里面是一个tuple，IP地址和PORT，发送命令给server，然后用result接收结果并进行return
