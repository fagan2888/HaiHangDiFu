import sqlite3
from datetime import date
from datetime import timedelta


class HHDF():
    """
    用于统计海航地服数据统计报表中的各项数据
    """
    def __init__(self):
        self.conn = sqlite3.connect(r"Data\data.db")
        self.cur = self.conn.cursor()

    @property
    def Nowadays(self):
        """返回昨天的日期"""
        value = date.today().strftime("%Y-%m-%d") + " 10:00"
        return value

    @property
    def one_days_ago(self):
        """返回一天前的日期"""
        value = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d") \
            + " 10:00"
        return value

    @property
    def two_days_ago(self):
        """返回两天前的日期"""
        value = (date.today() - timedelta(days=2)).strftime("%Y-%m-%d") \
            + " 10:00"
        return value

    @property
    def three_days_ago(self):
        """返回三天前的日期"""
        value = (date.today() - timedelta(days=3)).strftime("%Y-%m-%d") \
            + " 10:00"
        return value

    @property
    def start_month(self):
        """返回月份数据统计的起始时间"""
        value = date.today().strftime("%Y-%m") + "-01"
        return value

    @property
    def end_month(self):
        """返回月份数据统计的截至时间"""
        value = date.today().strftime("%Y-%m") + "-31"
        return value

    @property
    def start_year(self):
        """返回年份数据统计的起始时间"""
        value = date.today().strftime("%Y") + "-01-01"
        return value

    @property
    def end_year(self):
        """返回年份数据统计的截至时间"""
        value = date.today().strftime("%Y") + "-12-31"
        return value

    @property
    def one_days_ago_jiao_tong_sum(self):
        """返回前一天交通工具的保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.one_days_ago}' \
            AND [承保时间] < '{self.Nowadays}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def two_days_ago_jiao_tong_sum(self):
        """返回两天前交通工具的保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.two_days_ago}' \
            AND [承保时间] < '{self.one_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def three_days_ago_jiao_tong_sum(self):
        """返回三天前交通工具的保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.three_days_ago}' \
            AND [承保时间] < '{self.two_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def one_days_ago_jiao_tong_count(self):
        """返回前一天交通工具的件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.one_days_ago}' \
            AND [承保时间] < '{self.Nowadays}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def two_days_ago_jiao_tong_count(self):
        """返回两天前交通工具的件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.two_days_ago}' \
            AND [承保时间] < '{self.one_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def three_days_ago_jiao_tong_count(self):
        """返回三天前交通工具的件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.three_days_ago}' \
            AND [承保时间] < '{self.two_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def month_jiao_tong_sum(self):
        """返回交通工具的月保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_month}' \
            AND [承保时间] <= '{self.end_month}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def month_jiao_tong_count(self):
        """返回交通工具的月件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_month}' \
            AND [承保时间] <= '{self.end_month}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def year_jiao_tong_sum(self):
        """返回交通工具的年保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_year}' \
            AND [承保时间] <= '{self.end_year}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def year_jiao_tong_count(self):
        """返回交通工具的年件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_year}' \
            AND [承保时间] <= '{self.end_year}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '华安一路平安交通工具意外伤害保险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def one_days_ago_hang_kong_sum(self):
        """返回一天前航空意外的保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.one_days_ago}' \
            AND [承保时间] < '{self.Nowadays}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def two_days_ago_hang_kong_sum(self):
        """返回两天前航空意外的保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.two_days_ago}' \
            AND [承保时间] < '{self.one_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def three_days_ago_hang_kong_sum(self):
        """返回三天前航空意外的保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.three_days_ago}' \
            AND [承保时间] < '{self.two_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def one_days_ago_hang_kong_count(self):
        """返回一天前航空意外的件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.one_days_ago}' \
            AND [承保时间] < '{self.Nowadays}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def two_days_ago_hang_kong_count(self):
        """返回两天前航空意外的件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.two_days_ago}' \
            AND [承保时间] < '{self.one_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def three_days_ago_hang_kong_count(self):
        """返回三天前航空意外的件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.three_days_ago}' \
            AND [承保时间] < '{self.two_days_ago}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def month_hang_kong_sum(self):
        """返回航空意外的月保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_month}' \
            AND [承保时间] <= '{self.end_month}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def month_hang_kong_count(self):
        """返回航空意外的月件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_month}' \
            AND [承保时间] <= '{self.end_month}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def year_hang_kong_sum(self):
        """返回航空意外的年保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_year}' \
            AND [承保时间] <= '{self.end_year}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def year_hang_kong_count(self):
        """返回航空意外的年件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_year}' \
            AND [承保时间] <= '{self.end_year}' \
            AND [保单状态] <> '已退保' \
            AND [产品名称] = '航空意外险'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def one_days_ago_sum(self):
        """返回一天前的整体保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.one_days_ago}' \
            AND [承保时间] < '{self.Nowadays}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def two_days_ago_sum(self):
        """返回两天前的整体保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.two_days_ago}' \
            AND [承保时间] < '{self.one_days_ago}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def three_days_ago_sum(self):
        """返回三天前的整体保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.three_days_ago}' \
            AND [承保时间] < '{self.two_days_ago}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def one_days_ago_count(self):
        """返回一天前的整体件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.one_days_ago}' \
            AND [承保时间] < '{self.Nowadays}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def two_days_ago_count(self):
        """返回两天前的整体件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.two_days_ago}' \
            AND [承保时间] < '{self.one_days_ago}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def three_days_ago_count(self):
        """返回三天前的整体件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.three_days_ago}' \
            AND [承保时间] < '{self.two_days_ago}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def month_sum(self):
        """返回整体月保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_month}' \
            AND [承保时间] <= '{self.end_month}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def month_count(self):
        """返回整体月件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_month}' \
            AND [承保时间] <= '{self.end_month}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def year_sum(self):
        """返回整体年保费"""
        sql_str = f"SELECT SUM([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_year}' \
            AND [承保时间] <= '{self.end_year}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value

    @property
    def year_count(self):
        """返回整体年件数"""
        sql_str = f"SELECT COUNT([保费]) \
            FROM [保险清单] \
            WHERE  [承保时间] >= '{self.start_year}' \
            AND [承保时间] <= '{self.end_year}' \
            AND [保单状态] <> '已退保'"
        self.cur.execute(sql_str)
        value = self.cur.fetchone()[0]
        if value is None:
            value = 0
        return value
