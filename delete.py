import mysql.connector

# MYsql 연결 설정
connection = mysql.connector.connect(
    host = "localhost",      # MySql 서버주소를 넣을 예정
    user = "root",           # 사용자 이름
    password = "root",       # 비밀번호
    database = "python_test" # 사용할 데이터베이스
)

cursor = connection.cursor() # 데이터베이스 작업을 위한 객체 생성

sql = "delete from users where name = %s" # 특정 이름의 사용자 삭제
values = ("Encore", ) # 값이 하나더라도 string data_type은 들어갈 수 없다.

cursor.execute(sql, values)
connection.commit()

print(f"{cursor.rowcount}개의 행이 삭제되었습니다.")

cursor.close()
connection.close()