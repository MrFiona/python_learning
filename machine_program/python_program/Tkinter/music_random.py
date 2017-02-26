#!/usr/bin/env python 
#coding:utf-8
#就是找到硬盘中所有的MP3文件和wma文件，再随机打开其中的一个
import os,random  

disk=['C','D','E','F','G','H']   
def search_file(filename,search_path,pathsep=os.pathsep):  
        for path in search_path.split(pathsep):  
            candidate = os.path.join(path,filename)  
            if os.path.isfile(candidate):  
                return os.path.abspath(candidate)  
def random_play():  
        music=[]  
        found=False  
        for i in range(0,6):  
            for dirpath, dirnames, filenames in os.walk(disk[i]+':/'):  
                for filename in filenames:  
                    if os.path.splitext(filename)[1] == '.mp3':  
                        filepath = os.path.join(dirpath, filename)  
                        if os.path.getsize(filepath)>1048576:#这里是为了防止打开一些非音乐音频，比如游戏配音  
                            music.append(filepath)  
                            found=True  
        if not found:  
            for dirpath, dirnames, filenames in os.walk('C:/Users'):  
                for filename in filenames:  
                    if os.path.splitext(filename)[1] == '.mp3':  
                        filepath = os.path.join(dirpath, filename)  
                        if os.path.getsize(filepath)>1048576:  
                            music.append(filepath)  
                            found=True  
        if not found:  
            for i in range(0,5):  
                for dirpath, dirnames, filenames in os.walk(disk[i]+':/'):  
                    for filename in filenames:  
                        if os.path.splitext(filename)[1] == '.wma':  
                            filepath = os.path.join(dirpath, filename)  
                            if os.path.getsize(filepath)>1048576:  
                                music.append(filepath)  
                                found=True  
        if not found:  
            for dirpath, dirnames, filenames in os.walk('C:/Users'):  
                for filename in filenames:  
                    if os.path.splitext(filename)[1] == '.wma':  
                        filepath = os.path.join(dirpath, filename)  
                        if os.path.getsize(filepath)>1048576:  
                            music.append(filepath)  
                            found=True  
        random_music=random.choice(music)  
        os.startfile(random_music)  
        if not found:  
            print ('没有找到音频文件')
if __name__=='__main__':  
        random_play()  
