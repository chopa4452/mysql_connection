import mysql.connector

with mysql.connector.connect(
    host = "localhost",
    user = "root",
    password = "root",
    database = "python_test"
) as connection:
    
    with connection.cursor() as cursor:
        cursor.execute("select * from users")
        for row in cursor.fetchall():   # rows를 가져오는거랑 같은 문법
            print(row)
