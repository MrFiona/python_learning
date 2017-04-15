#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-14 12:09
# Author  : MrFiona
# File    : cache_mechanism.py
# Software: PyCharm Community Edition


import os
import re
import zlib
import pickle
from datetime import datetime, timedelta
import urlparse

class DiskCache(object):
    def __init__(self, cache_dir='cache', expires=timedelta(days=30)):
        self.cache_dir = cache_dir
        self.expires = expires

    def has_expired(self, timestamp):
        return datetime.utcnow() > timestamp + self.expires

    def url_to_path(self, url):
        components = urlparse.urlsplit(url)
        path = components.path
        # print path
        filename = components.netloc + path + components.query
        filename = re.sub('[^/0-9a-zA-Z\-.,;_ ]', '_', filename)
        filename = '/'.join(segment[:255] for segment in filename.split('/'))
        # print filename
        return_path = os.path.join(self.cache_dir, filename)
        return return_path

    def __getitem__(self, url):
        path = self.url_to_path(url)
        if os.path.exists(path):
            with open(path, 'rb') as fp:
                return pickle.load(fp)
        else:
            raise KeyError(url + ' does not exist')

    def __setitem__(self, url, result):
        path = self.url_to_path(url)
        folder = os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, 'wb') as fp:
            fp.write(pickle.dumps(result))

if __name__ == '__main__':
    cache = DiskCache()
    # cache.url_to_path('https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/BKC/2017%20WW13/5772_BKC.html')
    rr = cache['https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/BKC/2017%20WW13/5774_BKC.html']
    print  cache
    print rr