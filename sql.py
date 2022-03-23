import pymysql

# 打开数据库连接
connection = pymysql.connect(host='localhost', user='root', password='1234', db='test', charset='utf8')
# 使用cursor()方法获取操作游标,之后就可以输入sql语句
cur = connection.cursor()
# 使用execute方法执行SQL语句

# print(count)


# 使用 fetchone() 方法获取一条数据
# res = cur.fetchall()

# fields = cur.description
# print(list(fields))
sql = "INSERT INTO userinfo(name,pwd) VALUES (%s, %s);"
data = [("Alex", 'a'), ("Egon", 'v'), ("Yuan", 'b')]
# try:
# 批量执行多条插入SQL语句
cur.executemany(sql, data)
# 提交事务
connection.commit()

# except Exception as e:
#     # 有异常，回滚事务
#     connection.rollback()

ret = cur.execute("SELECT * from userinfo;")
print(ret)
print(cur.fetchall())
connection.close()
