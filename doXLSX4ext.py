#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
# 用于“钉钉”审批类 XLSX 文件解析器
# ==================================
# 2017年10月10日@成都
#
#

from __future__ import unicode_literals

import sys
import mongodb_class
import xlrd
import time
from datetime import datetime

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

TableName = {
             u"项目简称,工作内容,可上传问题单,": 'operation_req',

             u"项目简称,项目编号,项目阶段,阶段,立项状态,立项,"
             u"服务类型,类型,工作方式,申请人数,工作地点,开始时间,"
             u"结束时间,具体工作内容,附件,": 'data_engineering_req',

             u"项目简称,项目编号,项目负责人,任务分流,开始时间,"
             u"结束时间,执行地址,任务内容,": 'task_req',

             u"公共服务,类型,内容,注册邮箱,": 'public_services_req',

             u"项目信息,项目简称,项目编号,项目负责人,运维负责人,"
             u"故障基本描述,故障等级,简要描述,影响范围,故障过程,问题分析,"
             u"解决方案,故障报告附件,": 'fault_repair_req',

             u"外出类型,出差类型,开始时间2,结束时间2,时长,"
             u"起止地点,外出事由,附件,备注,": 'trip_req',

             u"项目简称,项目编号,所属业务,业务,项目状态,状态,"
             u"立项状态,立项,服务类型,类型,工作地点,开始时间,结束时间,工作内容,"
             u"起始时间,结束时间,附件,": 'sulotion_req',

             u"项目简称,项目编号,项目状态,项目阶段,工作方式,"
             u"工作类型,开始时间,结束时间,工作地点,工作内容,备注,附件,": 'project implementation',

             u"项目简称,项目编号,所属业务,业务,项目状态,状态,"
             u"立项状态,立项,服务类型,类型,人数,工作地点,开始时间,结束时间,"
             u"具体工作内容,附件,": 'data_sulotion_req',

             u"项目简称,项目编号,事由,付款方式,转账信息,"
             u"收款方名称,开户行,账号,出差起止日期,金额,往返交通费,补助,住宿费,"
             u"其他费用,金额小计,金额小计(大写),附件,备注,": 'loan_req',

             u"项目简称,项目编号,事由,付款方式,转账信息,"
             u"收款方名称,开户行,账号,出差明细,出差起止日期,出差起止地点,往返交通费,"
             u"补助,住宿费,其他费用,金额小计,金额小计(大写),附件,备注,": 'reimbursement_req',

             u"项目简称,项目编号,项目负责人,明细,产品名称及版本,"
             u"货物类型,货物清单,货物用途,其他用途,交付场地,": 'pd_outgoing_req',

             u"项目简称,工作内容,支撑明细,支撑人员,"
             u"起始日期,工作量预估（人.天）,附件,": 'pj_development_req',

             u"项目简称,项目编号,项目负责人,资料名称及版本,"
             u"出库类型,出库清单,附件,出库用途,其他用途,": 'file_outgoing_req',
             }


class XlsxHandler:

    def __init__(self, pathname):
        self.data = xlrd.open_workbook(pathname)
        self.tables = self.data.sheets()
        self.table = self.tables[0]
        self.nrows = self.table.nrows

    def getSheetNumber(self):
        return len(self.tables)

    def setSheet(self, n):
        if n < len(self.tables):
            self.table = self.tables[n]
            self.nrows = self.table.nrows

    def getNrows(self):
        return self.nrows

    def getData(self, row, col):
        """
        获取xlsx记录单元的数据
        :param table: 数据源
        :param row: 行号
        :param col: 列号
        :return:
        """
        try:
            return self.table.row_values(row)[col]
        except:
            return None

    def getXlsxColStr(self):
        _col = self.getXlsxColName(self.table.ncols)
        _str = ""
        for _s in _col:
            _str += (_s + ',')
        return _str, self.table.ncols

    def getXlsxColName(self, nCol):
        _col = []
        for i in range(14, nCol):
            _colv = self.table.row_values(0)[i]
            _col.append(_colv)
        return _col

    def getXlsxAllColName(self, nCol):
        _col = []
        for i in range(nCol):
            _colv = self.table.row_values(0)[i]
            _col.append(_colv)
        return _col

    def getXlsxRow(self, i, nCol, lastRow):
        """
        获取某行数据
        :param i: 行号
        :param nCol: 字段数
        :param lastRow: 上一行数据
        :return: 指定行的数据
        """

        """ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
        """
        _row = []
        for _i in range(nCol):
            __row = self.table.row_values(i)[_i]
            __ctype = self.table.cell(i, _i).ctype
            if __ctype == 3:
                _date = datetime(*xlrd.xldate_as_tuple(__row, 0))
                __row = _date.strftime('%Y/%m/%d')
                __row = __row.split('/')
                __row = u"%d年%d月%d日" % (int(__row[0]), int(__row[1]), int(__row[2]))
            elif __ctype == 2:
                __row = str(__row)
            elif __ctype == 5:
                __row = ''
            _row.append(__row)

        row = []
        _i = 0
        for _r in _row:
            if _r is None or len(str(_r)) == 0:
                """用第一行的内容填充合并字段的其它单元"""
                if lastRow is not None:
                    _r = lastRow[_i]
            row.append(_r)
            _i = _i + 1
        return row


def doList(xlsx_handler, mongodb, _type, _op, _ncol, keys):

    _keys = keys
    _rows = []
    for i in range(1, xlsx_handler.getNrows()):
        _row = xlsx_handler.getXlsxRow(i, _ncol, None)
        _rows.append(_row)

    _col = xlsx_handler.getXlsxAllColName(_ncol)
    _count = 0
    _key = []
    if len(_rows) > 0:
        if 'APPEND' in _op:
            '''追加方式，如日志记录
            '''
            for _row in _rows:
                _value = {}
                _i = 0
                _search = {}
                for _v in _keys:
                    _search[_col[_v]] = _row[_v]
                for _c in _col:
                    _value[_c] = _row[_i]
                    _i += 1
                try:
                    if _search not in _key:
                        mongodb.handler(_type, 'update', _search, _value)
                        _key.append(_search)
                        _count += 1
                except Exception, e:
                    print "error: ", e
                finally:
                    print '.',
    print "[", _count, "]"
    return {"OK": True, "INFO": u"导入%d条数据" % _count}


def main(filename):

    mongo_db = mongodb_class.mongoDB('ext_system')

    if filename is None:
        filename = sys.argv[1]
    print filename

    xlsx_handler = XlsxHandler(filename)

    try:
        _str, _ncols = xlsx_handler.getXlsxColStr()
        if _str not in TableName:
            # print(">>> Err: [%s][%s] not be recognised" % (filename, _str))
            return {"OK": False, "INFO": u">>> Err: [%s][%s] not be recognised" % (filename, _str)}
        _table = TableName[_str].lower()
        _ret = doList(xlsx_handler, mongo_db, _table, "APPEND", _ncols, range(3))
        # print("%s- Done" % time.ctime())
        return _ret

    except Exception, e:
        # print e
        return {"OK": False, "INFO": "%s" % e}


if __name__ == '__main__':
    main(None)

#
# Eof
