# -*- coding: utf-8 -*-
#
#   生成项目周报，并通过邮件发送
#   ============================
#   2019.8.1 @Chengdu
#

from datetime import datetime, date, timedelta
import send_email_file
import sys
import PjWeeklyRpt

Emails = {
    "To": "lixiaowei@chinacloud.com.cn,xiaoqingshan@chinacloud.com.cn,"
          "wuyumin@chinacloud.com.cn,xiangxiaoyan@chinacloud.com.cn,"
          "guchenchen@chinacloud.com.cn",
    "Cc": "zhangjing_sh@chinacloud.com.cn,shenwei@chinacloud.com.cn"
}


def write_title(_book, _titles):
    _v = 0
    for _t in _titles:
        _book.write((0, _v, _t))
        _v += 1


def main():
    """
    主程序
    :return:
    """

    """获得上一周的周一和周日"""
    _n = date.today().weekday()
    _monday = datetime.now() - timedelta(days=(_n+7))
    _sunday = datetime.now() - timedelta(days=_n+1)

    _bg_date = "%02d%02d%02d" % (_monday.year, _monday.month, _monday.day)
    _ed_date = "%02d%02d%02d" % (_sunday.year, _sunday.month, _sunday.day)

    print _bg_date, _ed_date

    PjWeeklyRpt.main(_bg_date, _ed_date)
    _mail = {
        "Smtp_Server": "smtp.chinacloud.com.cn",
        "Smtp_Password": sys.argv[1],
        "Receivers": Emails["To"],
        "Cc": Emails["Cc"],
        "From": "RD-MIS@chinacloud.com.cn",
        "To": Emails["To"],
        "Msg_Title": "An Auto-send email by R&D MIS ",
        "Smtp_Sender": "shenwei@chinacloud.com.cn",
        "Text": "由R&D MIS系统自动生成的<%s %s>项目周报见附件，请参考。" % (_bg_date, _ed_date),
        "Files": ["weekly_rpt.docx"]
    }
    print u"%s" % _mail["Text"], _mail["Files"]
    send_email_file.EmailClass(_mail).send()


if __name__ == '__main__':
    main()

