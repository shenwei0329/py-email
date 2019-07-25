# -*- coding: utf-8 -*-

import sys
import doXLSX4ext_file

# from __future__ import unicode_literals

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')


def file_handler(_file):

    if (('.xlsx' in _file) or ('.xls' in _file)) and ('~$' not in _file):
        return doXLSX4ext_file.main(_file)
    else:
        # print "Invalid file: ", _file
        return {"OK": False, "INFO": "invalid file name"}

