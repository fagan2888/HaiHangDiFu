# _*_ coding: utf-8 _*_

import sqlite3
import logging
import xlrd

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s | %(message)s")


def update():
    conn = sqlite3.connect(r"Data\data.db")
    cur = conn.cursor()
    logging.debug("数据库连接成功")

    book = xlrd.open_workbook(r"保单信息.xlsx")
    sh = book.sheet_by_name("保单信息")
    logging.debug("Excel文件导入成功")

    logging.debug(f"需要导入{sh.nrows-5}条数据")
    i = 3

    while i <= sh.nrows - 3:
        sql_str = f"INSERT INTO [保险清单] VALUES ('{sh.row_values(i)[0]}', \
    '{sh.row_values(i)[1]}', \
    {sh.row_values(i)[2]}, \
    '{sh.row_values(i)[3]}', \
    '{sh.row_values(i)[4]}', \
    '{sh.row_values(i)[5]}', \
    '{sh.row_values(i)[6]}', \
    {sh.row_values(i)[7]}, \
    '{sh.row_values(i)[8]}', \
    '{sh.row_values(i)[9]}', \
    '{sh.row_values(i)[10]}', \
    '{sh.row_values(i)[11]}', \
    '{sh.row_values(i)[12]}', \
    '{sh.row_values(i)[13]}', \
    '{sh.row_values(i)[14]}', \
    '{sh.row_values(i)[15]}', \
    '{sh.row_values(i)[16]}', \
    '{sh.row_values(i)[17]}', \
    '{sh.row_values(i)[18]}')"
        cur.execute(sql_str)

        i += 1
        if (i - 4) % 100 == 0:
            conn.commit()
            logging.debug(f"已导入{i-4} / {sh.nrows-5}条数据")

    conn.commit()

    logging.debug("数据库更新完成")
    cur.close()
    conn.close()
    logging.debug("数据库关闭完成")
