# -*- coding: utf-8 -*-

import sys
import doXLSX4Duty
import json
import locale

# from __future__ import unicode_literals

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')


def file_handler(_file):

    f = open('duty_file.txt', 'r')
    file_list = f.read()
    f.close()

    _short_file = _file.split("\\")[-1]

    if _short_file in file_list:
        print "Nothing to do!"
        return

    print "fileHandler: ", _short_file

    if ('.xlsx' in _short_file) and\
            ('~$' not in _short_file) and\
            (u'考勤报表'.encode(locale.getdefaultlocale()[1]) in _short_file):
        _ret = doXLSX4Duty.main(_file)
        file_list += _short_file
    else:
        # print "Invalid file name: ", _short_file
        return {"OK": False, "INFO": "invalid file name"}

    f = open('duty_file.txt', 'w')
    f.write(file_list)
    f.close()

    return _ret

