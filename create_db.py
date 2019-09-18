# _*_ coding: utf-8 _*_

import sqlite3
import logging


logging.basicConfig(level=logging.DEBUG, format="%(asctime)s | %(message)s")

conn = sqlite3.connect(r"Data\data.db")
cur = conn.cursor()
logging.debug("数据库连接成功")

sql_str = "CREATE TABLE [保险清单] ([保单号] TEXT, [产品名称] TEXT, [保费] REAL, \
    [支付方式] TEXT, [保单状态] TEXT, [退保来源] TEXT, [被保险人名称] TEXT, \
    [凭证打印次数] INTEGER, [推广工号] TEXT, [推广业务员] TEXT, [承保时间] TEXT, \
    [起保时间] TEXT, [保单失效时间] TEXT, [保险期限] TEXT, [销售机场] TEXT, [销售来源] TEXT, \
    [业务员工号] TEXT, [业务员名称] TEXT, [承保机构] TEXT)"

cur.execute(sql_str)

conn.commit()

cur.close()
conn.close()
logging.debug("数据库关闭完成")
