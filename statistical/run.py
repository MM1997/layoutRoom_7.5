from statistical.readMysql import Mysql
from statistical.excle import Excle


if __name__ == "__main__":
    path = r'statistical.xlsx'
    mysql = Mysql()
    datas = []
    sr_key_xs= [0.75, 0.85, 0.95, 1]
    for sr_key_x in sr_key_xs:
        sql = 'select room_name,solutionId,dnaSolutionId ,count(1) from data.layout_best_infos where 1=1 ' \
              'and sr_key_x>0 ' \
              'and sr_key_x<{} ' \
              'and room_name in ("客厅","餐厅","主卧","次卧","儿童房","老人房","书房","榻榻米房") ' \
              'group by room_name,solutionId ' \
              'order by room_name,solutionId limit 100'.format(sr_key_x)

        data = mysql.get_data(sql)
        datas.append(data)

    excle = Excle(path)
    excle.statistical(datas, keys=["room_name"], values=["count(1)"], conditions=sr_key_xs)




