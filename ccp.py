import pymysql
import os
import csv
import pandas

class Mysql():
    def __init__(self,host,user,password):
        self.host = host
        self.user = user
        self.password = password
        self.connection = pymysql.connect(host=self.host,user=self.user,password=self.password,charset='utf8')
        self.cursor = self.connection.cursor()
    #添加表单
    def UpdateTerm(self,name,rows):
        self.cursor.execute('use Activity')
        try:
            self.cursor.execute('drop table `'+name+"`")
        except (Exception):
            print(">表单不存在，直接创建新表单")
        else:
            print('>表单存在，重新创建表单')
        #创建表单
        sql = 'create table `'+name+"` ("
        for term in rows[0]:
            if term == '学号':
                term = 'id'
            elif term == '姓名':
                term = 'name'
            elif term == '总时长':
                term = 'hour'
            else:
                term = "`" + str(term) + "`"
            sql += str(term) + " varchar(255),"
        sql = sql[:-1]
        sql = sql + ')'
        print(sql)
        self.cursor.execute(sql)
        #开始插入列表
        flag = False
        for row in rows:
            if flag==False:
                flag=True
                sql = "insert into `"+name+"` values"
            else:
                sql +='('
                for term in row:
                    sql += "'"+term+ "'"+ ','
                sql = sql[:-1]
                sql += '),'
        sql = sql[:-1]
        try:
            self.cursor.execute(sql)
        except Exception:
            print(">数据插入错误！")
        else:
            print(">数据导入成功！")
        self.connection.commit()

def Excel2CSV(filename):
    data_xls = pandas.read_excel(filename, 'Sheet1', index_col=None)  
    data_xls.to_csv('update.csv', encoding='utf-8', index=None)

if __name__ == '__main__':
    filename = input(">Welcome to CCP!\n>Please enter filename:")
    Excel2CSV(filename)
    name = input(">Please enter semester:")
    csv_reader = csv.reader(open('update.csv', encoding='utf-8'))
    rows = [row for row in csv_reader]
    mysql = Mysql('ibiter.org','root','wang@85#2')
    mysql.UpdateTerm(name,rows)
    input(">Press any key to exit!")