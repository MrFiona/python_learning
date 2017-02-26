#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import urllib2
import os
import threading
import socket
import httplib
import sys
reload(sys)
import time
import os.path
import chardet

sys.setdefaultencoding('utf-8')
# year_list = ['2009', '2010', '2011', '2012', '2013', '2014', '2015']
year_list = ['2015']
# year_list = ['2009']
# year_list = list(reversed(year_list))
url_prefix = 'http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/'
mode_flag = True
mode1_flag = True

def thread_func_spider():
    global  url_prefix, year_list, mode1_flag, mode_flag
    for link in xrange(len(year_list)):
        try:
            m = urllib2.urlopen(url_prefix + '{}/index.html'.format(year_list[link])).read()
            decode_method = chardet.detect(m)['encoding']
            print 'province:\t%s' %(decode_method)
            content = m.decode(decode_method)

            pattern_1 = re.compile(r"<a href='(\d+).html'>(.+?)<br/></a></td>")
            pattern_2 = re.compile(r"<a href='(\d+).html'>(.+?)</a></td>")

            match_1 = pattern_1.findall(content)
            if not match_1:
                match_1 = pattern_2.findall(content)
            province_num_list = list(match_1[demo_province][0] for demo_province in range(len(match_1)))
            province_name_list = list(match_1[demo_province][1] for demo_province in range(len(match_1)))
            print os.getcwd() + '\n'
            print os.path.exists(r'{0}_{1}\n' .format(int(year_list[link]),u'year_province'))
            if not os.path.exists(r'{0}_{1}' .format(int(year_list[link]),u'year_province')):
                try:
                    os.mkdir(r'C:\Users\apple\PycharmProjects\personal_program\test\{0}_{1}'.format(int(year_list[link]),u'year_province'))
                except WindowsError as e:
                    print e.message

            os.chdir(r'C:\Users\apple\PycharmProjects\personal_program\test\{0}_{1}'.format(int(year_list[link]),u'year_province'))
            # 将爬取的数据存入文件中
            for province in xrange(len(province_name_list)):
                # 将省份信息写进文件中
                with open('{0}_{1}.txt'.format(int(year_list[link]), u'year_province'), 'a+') as file:
                    if '<br/>' in province_name_list[province]:
                        province_name_list[province] = province_name_list[province].replace('<br/>', '')
                    file.write(province_name_list[province] + '\n')

            # 获取省份下的市的信息
            for city in xrange(len(province_num_list)):
                try:
                    m = urllib2.urlopen(url_prefix + '{0}/{1}.html'.format(year_list[link], province_num_list[city])).read()
                    decode_method = chardet.detect(m)['encoding']
                    print 'city:\t%s' % (decode_method)
                    content_city = m.decode(decode_method)
                    pattern_city = re.compile(
                        r"<a href='(.+?).html'>(\d+)</a></td><td><a href='{0}/(\d+).html'>(.+?)</a></td></tr>".format(
                            int(province_num_list[city])))
                    match_city = pattern_city.findall(content_city)
                    # print match_city
                    city_url_list = list(match_city[demo_city][0] for demo_city in range(len(match_city)))
                    city_name_list = list(match_city[demo_city][1] for demo_city in range(len(match_city)))
                    city_num_list = list(match_city[demo_city][2] for demo_city in range(len(match_city)))
                    city_code_list = list(match_city[demo_city][3] for demo_city in range(len(match_city)))
                    os.chdir(r'C:\Users\apple\PycharmProjects\personal_program\test\{0}_{1}'.format(int(year_list[link]),u'year_province'))
                    print os.path.exists(r'{0}_{1}'.format(int(year_list[link]), int(province_num_list[city])))
                    print
                    if not os.path.exists(r'{0}_{1}'.format(int(year_list[link]), int(province_num_list[city]))):
                        try:
                            os.mkdir(os.getcwd() + os.sep + '{0}_{1}'.format(int(year_list[link]), int(province_num_list[city])))
                        except WindowsError as e:
                            print e.message
                    os.chdir(os.getcwd() + os.sep + '{0}_{1}'.format(int(year_list[link]), int(province_num_list[city])))
                    with open('{0}_{1}.txt'.format(int(year_list[link]), int(province_num_list[city])), 'a+') as file:
                        for i in xrange(len(city_name_list)):
                            if mode_flag:
                                # file.write('代码\t\t\t\t名称\n')
                                mode_flag = False
                            file.write('{:>12}{:>24}\n'.format(city_name_list[i], city_code_list[i]))
                    # 获取市下的区县信息
                    for county in xrange(len(city_num_list)):
                        # m = year_list[link]
                        # n = province_num_list[city]
                        # p = city_num_list[county]
                        try:
                            m = urllib2.urlopen(url_prefix + '{0}/{1}.html'.format(year_list[link],city_url_list[county])).read()
                            decode_method = chardet.detect(m)['encoding']
                            print 'street:\t%s' % (decode_method)
                            content_county = m.decode(decode_method)
                            pattern_county = re.compile(
                                r"<a href='(.+?).html'>(\d+)</a></td><td><a href='(\d+)/(\d+).html'>(.+?)</a></td></tr>")
                            match_county = pattern_county.findall(content_county)
                            county_url_list = list(
                                match_county[demo_county][0] for demo_county in range(len(match_county)))
                            county_name_list = list(
                                match_county[demo_county][1] for demo_county in range(len(match_county)))
                            county_unicode_list = list(
                                match_county[demo_county][2] for demo_county in range(len(match_county)))
                            county_num_list = list(
                                match_county[demo_county][3] for demo_county in range(len(match_county)))
                            county_code_list = list(
                                match_county[demo_county][4] for demo_county in range(len(match_county)))
                            os.chdir(r'C:\Users\apple\PycharmProjects\personal_program\test\{0}_year_province\{0}_{1}'.format(int(year_list[link]), int(province_num_list[city])))
                            if not os.path.exists(r'{0}_{1}_{2}' .format(int(year_list[link]),int(province_num_list[city]),int(city_num_list[county]))):
                                try:
                                    os.mkdir(os.getcwd() + os.sep + '{0}_{1}_{2}'.format(int(year_list[link]),int(province_num_list[city]),int(city_num_list[county])))
                                except WindowsError as e:
                                    print e.message
                            os.chdir(os.getcwd() + os.sep + '{0}_{1}_{2}'.format(int(year_list[link]),int(province_num_list[city]),int(city_num_list[county])))
                            with open('{0}_{1}_{2}.txt'.format(int(year_list[link]), int(province_num_list[city]),int(city_num_list[county])), 'a+') as file:
                                for i in xrange(len(county_name_list)):
                                    if mode1_flag:
                                        # file.write('代码\t\t\t\t名称\n')
                                        mode1_flag = False
                                    file.write('{:>12}{:>24}\n'.format(county_name_list[i], county_code_list[i]))
                            # 获取区县下的乡镇信息
                            for street in xrange(len(county_num_list)):
                                m = county_url_list[street]
                                try:
                                    m = urllib2.urlopen(url_prefix + '{0}/{1}/{2}.html'.format(int(year_list[link]),int(province_num_list[city]),county_url_list[street])).read()
                                    decode_method = chardet.detect(m)['encoding']
                                    print 'towns:\t%s' % (decode_method)
                                    content_street = m.decode(decode_method)
                                    pattern_street = re.compile(
                                        r"<a href='(.+?).html'>(\d+)</a></td><td><a href='(\d+)/(\d+).html'>(.+?)</a></td></tr>")
                                    match_street = pattern_street.findall(content_street)
                                    street_url_list = list(
                                        match_street[demo_street][0] for demo_street in range(len(match_street)))
                                    street_name_list = list(
                                        match_street[demo_street][1] for demo_street in range(len(match_street)))

                                    street_num_list = list(
                                        match_street[demo_street][3] for demo_street in range(len(match_street)))
                                    street_code_list = list(
                                        match_street[demo_street][4] for demo_street in range(len(match_street)))
                                    # os.chdir(r'C:\Users\apple\PycharmProjects\personal_program\test\{0}_year_province\{0}_{1}\{0}_{1}_{2}'.format(int(year_list[link]), int(province_num_list[city]), int(city_num_list[county])))
                                    # os.mkdir(os.getcwd() + os.sep + '{0}_{1}_{2}_town'.format(int(year_list[link]), int(province_num_list[city]),int(city_num_list[county])))
                                    # os.chdir(os.getcwd() + os.sep + '{0}_{1}_{2}_town'.format(int(year_list[link]), int(province_num_list[city]),int(city_num_list[county])))
                                    with open(
                                            '{0}_{1}_{2}_22.txt'.format(int(year_list[link]),
                                                                        int(province_num_list[city]),
                                                                        int(city_num_list[county])), 'a+') as file:
                                        for i in xrange(len(street_name_list)):
                                            if mode1_flag:
                                                # file.write('代码\t\t\t\t名称\n')
                                                mode1_flag = False
                                            file.write(
                                                '{:>12}{:>24}\n'.format(street_name_list[i], street_code_list[i]))
                                    # 获取村乡镇下的（居委会）信息
                                    for village in xrange(len(street_num_list)):
                                        if not street_url_list[village] or not county_unicode_list[street] or not \
                                        year_list[
                                            link] or not province_num_list[city]:
                                            continue
                                        try:
                                            m = urllib2.urlopen(url_prefix + '{0}/{1}/{2}/{3}.html'.format(int(year_list[link]),int(province_num_list[city]),county_unicode_list[street],street_url_list[village])).read()
                                            decode_method = chardet.detect(m)['encoding']
                                            print 'village:\t%s' % (decode_method)
                                            content_village = m.decode(decode_method)
                                            pattern_village = re.compile(
                                                r"<tr class='villagetr'><td>(\d+)</td><td>(\d+)</td><td>(.+?)</td></tr>")
                                            match_village = pattern_village.findall(content_village)
                                            village_code_list = list(
                                                match_village[demo_village][0] for demo_village in
                                                range(len(match_village)))
                                            village_classify_list = list(
                                                match_village[demo_village][1] for demo_village in
                                                range(len(match_village)))
                                            village_name_list = list(
                                                match_village[demo_village][2] for demo_village in
                                                range(len(match_village)))
                                            # os.chdir(
                                            #     r'C:\Users\apple\PycharmProjects\personal_program\test\{0}_year_province\{0}_{1}\{0}_{1}_{2}\{0}_{1}_{2}_{3}'.format(
                                            #         int(year_list[link]), int(province_num_list[city]), int(city_num_list[county]), int(street_num_list[street])))
                                            # os.mkdir(os.getcwd() + os.sep + '{0}_{1}_{2}_{3}_{4}'.format(int(year_list[link]),int(province_num_list[city]),int(city_num_list[county]),
                                            #                                                          int(street_num_list[street]), int(village_code_list[village])))
                                            # os.chdir(os.getcwd() + os.sep + '{0}_{1}_{2}_{3}_{4}'.format(int(year_list[link]),int(province_num_list[city]),int(city_num_list[county]),
                                            #                                                          int(street_num_list[street]), int(village_code_list[village])))
                                            with open('{0}_{1}_{2}_street_1.txt'.format(int(year_list[link]),
                                                                                        int(province_num_list[city]),
                                                                                        int(city_num_list[county])),
                                                      'a+') as file:
                                                for i in xrange(len(village_name_list)):
                                                    if mode1_flag:
                                                        file.write('代码\t\t\t\t城乡分类\t\t\t\t名称\n')
                                                        mode1_flag = False
                                                    file.write(
                                                        '{:>12}{:>24}{:>24}\n'.format(village_code_list[i],
                                                                                             village_classify_list[
                                                                                                 i],
                                                                                             village_name_list[i]))
                                        except socket.timeout, e:
                                            pass
                                        except urllib2.URLError, ee:
                                            pass
                                        except httplib.BadStatusLine:
                                            pass
                                        except UnicodeDecodeError as eee:
                                            print eee.message
                                except socket.timeout, e:
                                    pass
                                except urllib2.URLError, ee:
                                    pass
                                except httplib.BadStatusLine:
                                    pass
                                except UnicodeDecodeError as eee:
                                    print eee.message
                        except socket.timeout, e:
                            pass
                        except urllib2.URLError, ee:
                            pass
                        except httplib.BadStatusLine:
                            pass
                        except UnicodeDecodeError as eee:
                            print eee.message
                except socket.timeout, e:
                    pass
                except urllib2.URLError, ee:
                    pass
                except httplib.BadStatusLine:
                    pass
                except UnicodeDecodeError as eee:
                    print eee.message
        except socket.timeout, e:
            pass
        except urllib2.URLError, ee:
            pass
        except httplib.BadStatusLine:
            pass
        except UnicodeDecodeError as eee:
            print eee.message






# def Using_Multithing_Fetch_Data(url_prefix, year_list, mode_flag=True, mode1_flag=True):
# 爬取省份以及获取省份接口


if __name__ == '__main__':
    t1 = time.clock()
    # thread_func_spider()

    thread_list = []
    thread_count = 1
    try:
        for link in xrange(thread_count):
            thread_list.append(threading.Thread(target=thread_func_spider))
            # time.sleep(10)
        for th in thread_list:
            th.start()
            print th
        # time.sleep(100)
        print thread_list[0].isAlive
        for th in thread_list:
            th.join()
    except WindowsError as e:
        print e.message
    t2 = time.clock()
    print t2 - t1


# for th in thread_list:
#     th.join()
# Using_Multithing_Fetch_Data(url_prefix, year_list, mode_flag, mode1_flag)
# t = threading.Thread(target=thread_func_spider, args=(url_prefix, year_list, mode_flag, mode1_flag))
# t.start()
# t.join()
# thread_func_spider(url_prefix, year_list, mode_flag, mode1_flag)

# Using_Multithing_Fetch_Data(url_prefix, year_list, mode_flag, mode1_flag)
#多线程同时进行不同年份爬取
# thread_list = []
# for thread_t in range(len(year_list)):
#     t = threading.Thread(target=Using_Multithing_Fetch_Data, args=(url_prefix, year_list, mode_flag, mode1_flag))
# t.start()
# t.join()
