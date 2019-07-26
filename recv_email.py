#-*- coding: UTF-8 -*-
#
#   接收邮件
#   ========
#   2019-07-23 Created by shenwei @Chengdu
#   从指定邮箱上取出从昨天到现在的邮件及其附件
#
#

from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import poplib
import sys
from datetime import datetime, date, timedelta
import pm_daily
import star_task
import duty
import send_email
import hashlib

rx_msg_list = []


def parser_date(msg):
    """
    解析邮件创建日期
    :param msg: 邮件
    :return: 创建日期（yyyymmdd）
    """
    _date = msg.get('date').split(',')[1][1:]
    if ',' in _date:
        _date = _date.split(', ')[1]
    utcstr = _date.replace('+00:00','').split(" (")[0]

    # print utcstr
    try:
        utcdatetime = datetime.strptime(utcstr, '%d %b %Y %H:%M:%S +0000 (GMT)')
    except:
        print utcstr
        utcdatetime = datetime.strptime(utcstr, '%d %b %Y %H:%M:%S +0800')
    ti = utcdatetime
    return "%04d%02d%02d" % (ti.year, ti.month, ti.day), utcstr


def parser_subject(msg):
    """
    解析邮件主题
    :param msg: 邮件
    :return: 主题
    """
    subject = msg['Subject']
    value, charset = decode_header(subject)[0]
    # print value, charset
    if charset:
        value = value.decode(charset)
    # print(u'subject: {0}'.format(value))
    return value
 

def parser_address(msg):
    """
    解析邮件地址
    :param msg: 邮件
    :return: 邮件地址
    """
    hdr, addr = parseaddr(msg['From'])
    name, charset = decode_header(hdr)[0]
    if charset:
        name = name.decode(charset)
    # print('sender: {0}，email: {1}'.format(name, addr))
    return name, addr


def parser_content(msg, _cr_date, _utc_date, _sender, _subject):
    """
    解析邮件正文
    :param msg: 邮件
    :param _cr_date: 邮件创建日期（用于文件命名）
    :param _sender: 发信人（用于文件命名）
    :return:
    """
    global rx_msg_list

    for par in msg.walk():
        name = par.get_param("name")
        if name:
            """具有附件文件"""
            value, charset = decode_header(name)[0]
            if charset:
                value = value.decode(charset)
            f_name = value
            print(">>> %s <<<" % f_name)
            data = par.get_payload(decode=True)
            """以二进制方式写入数据"""
            try:
                f = open("files/%s" % f_name, "wb")
            except:
                f = open("files/%s-%s-ref" % (_cr_date, _sender), "wb")
            f.write(data)
            f.close()

            if "chinacloud.com.cn" in _sender:
                """是公司内部邮件"""
                _ret = None
                if u"项目日报" in f_name and "doc" in f_name:
                    print "pm_daily.file_handler:",f_name
                    _ret = pm_daily.file_handler("files/%s" % f_name)
                elif _subject in [u"代码提交", u"项目经理", u"需求", u"北区运维"]:
                    _ret = star_task.file_handler("files/%s" % f_name)
                elif _subject in [u"考勤"]:
                    _ret = duty.file_handler("files/%s" % f_name)
                else:
                    print u"Subject: %s" % _subject
                if _ret is not None:
                    """定义邮件回复信息"""
                    rx_msg_list.append({"sender": _sender,
                                        "subject": _subject,
                                        "date": _utc_date,
                                        "file": f_name,
                                        "ret": _ret})
        else:
            _datas = par.get_payload(decode=True)
            if _datas is None:
                continue
            """正文文件命名"""
            f = open("htmls/%s-%s.html" % (_cr_date, _sender), "wb")
            for _data in _datas:
                f.write(_data)
            f.close()


"""扫描昨天到今天的邮件"""
yesterday = date.today() + timedelta(days=-1)
_arg_date = yesterday.strftime("%Y%m%d")

email = "shenwei@chinacloud.com.cn"
password = sys.argv[1]
pop3_server = "pop.chinacloud.com.cn"
 
server = poplib.POP3_SSL(pop3_server)
server.set_debuglevel(0)

server.user(email)
server.pass_(password)
 
# print("信息数量：%s 占用空间 %s" % server.stat())
resp, mails, octets = server.list()

_index = len(mails)
_err = False

f = open('email_hash.txt', 'r')
hash_list = f.read()
f.close()
_need_rewrite = False

for _idx in range(_index, 0, -1):

    # print(u"第%d封邮件：\n" % _idx)
    # print("="*80)

    _err = False
    try:
        resp, lines, ocetes = server.retr(_idx)
    except Exception, e:
        print(">>>Err: %s" % e)
        _err = True

        """当出现server.retr错误时，需要释放它，并重新创建一个"""
        server.quit()
        server = poplib.POP3_SSL(pop3_server)
        server.user(email)
        server.pass_(password)
    finally:
        if not _err:
            msg_content = b"\r\n".join(lines).decode("gbk")
            try:
                msg = Parser().parsestr(msg_content)
            except Exception, e:
                print(">>>Err.parsestr: %s" % e)
                _err = True
            finally:
                if not _err:
                    _name, _sender = parser_address(msg)
                    _subject = parser_subject(msg)
                    _date, _utc_date = parser_date(msg)
                    if _date >= _arg_date:
                        """收取正文和附件"""
                        _text = u"%s-%s-%s-%s" % (_name, _sender, _subject, _utc_date)
                        _hex = hashlib.md5(_text.encode("utf-8")).hexdigest()
                        if _hex not in hash_list:
                            parser_content(msg, _date, _utc_date, _sender, _subject)
                            hash_list += "%s\n" % _hex
                            _need_rewrite = True
                    else:
                        break

    print("-"*80)

# 关闭连接
server.quit()

if _need_rewrite:
    f = open('email_hash.txt', 'w')
    f.write(hash_list)
    f.close()

for _msg in rx_msg_list:

    if "ret" in _msg:
        if _msg["ret"]["OK"]:
            _text = "R&D MIS系统自动导入邮件 <%s>@<%s> 的附件<%s>数据，处理结果：<%s>" % (
                _msg["subject"], _msg["date"], _msg["file"], _msg["ret"]["INFO"])
        else:
            _text = "R&D MIS系统<错误提示>：邮件 <%s>@<%s> 的附件<%s>数据未被导入，原因：<%s>" % (
                _msg["subject"], _msg["date"], _msg["file"], _msg["ret"]["INFO"])

        mail = {
            "Smtp_Server": "smtp.chinacloud.com.cn",
            "Smtp_Password": sys.argv[1],
            "Receivers": [_msg["sender"], "shenwei@chinacloud.com.cn"],
            "Msg_Title": "An Auto-Reply email by R&D MIS-",
            "Smtp_Sender": "shenwei@chinacloud.com.cn",
            "From": "RD-MIS@chinacloud.com.cn",
            "To": _msg["sender"],
            "Text": _text
        }
        send_email.EmailClass(mail).send()
        print "need to send it!"


#
