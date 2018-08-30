#coding: utf-8

import numpy as np
import pymysql
import pandas as pd
import openpyxl
from openpyxl.styles import Font, colors, Alignment,Border,Side
from openpyxl.utils import get_column_letter,column_index_from_string
from pandas import DataFrame


class Mysql(object):
    def __init__(self,host='192.168.1.51',port=3308,user='admin',passwd='admin',db='data',charset='utf8'):
        self.host = host
        self.port = port
        self.user = user
        self.passwd = passwd
        self.db = db
        self.charset = charset
        self.config = {'host': self.host,
                  'port': self.port,
                  'user': self.user,
                  'passwd': self.passwd,
                  'db': self.db,
                  'charset': self.charset}
        #连接数据库
        self.conn = pymysql.connect(**self.config)

    def get_data(self,sql):
        '''
        将查询到的数据库数据转化为字典格式,如：
        [{'room_name': '主卧', 'solutionId': 11410, 'dnaSolutionId': 9324, 'count(1)': 207}, {'room_name': '书房', 'solutionId': 10589, 'dnaSolutionId': 5771, 'count(1)': 80}]
        :param sql:sql语句
        :return:
        '''
        df = pd.read_sql(sql,self.conn)
        info = df.to_dict('records')
        print(info)
        return info




sql = 'select room_name,solutionId,dnaSolutionId ,count(1) from data.layout_best_infos where 1=1 and sr_key_x>0 and sr_key_x<0.85 and room_name in ("客厅","餐厅","主卧","次卧","儿童房","老人房","书房","榻榻米房") group by room_name order by room_name limit 100'
# mysql = Mysql()
# data = mysql.get_data(sql)



def generate_new_dict(infos,keys=list,values = "count"):
    """
    :param infos: 字典:[{'room_name': '主卧', 'solutionId': 3445, 'count': 183}, {'room_name': '主卧', 'solutionId': 3822, 'count': 183}]
    :param keys: 生成要组成的key：list = ["room_name","solutionId"]
    :param values:
    :return: {'主卧-11410': 207, '书房-10589': 80, '儿童房-11422': 113, '客厅-11412': 407, '榻榻米房-6289': 42, '次卧-11008': 327, '老人房-8304': 42}
    """
    count = {}
    for info in infos:
        a = []
        for i in keys:
            a.append(str(info[i]))
        #生成要组成的key值
        key = "-".join(a)

        if key in count:
            count[key] = count[key] + info[values]
        else:
            count.setdefault(key, info[values])
    print(count)
    return count
