import mysql.connector

# MYsql 연결 설정
connection = mysql.connector.connect(
    host = "localhost",      # MySql 서버주소를 넣을 예정
    user = "root",           # 사용자 이름
    password = "root",       # 비밀번호
    database = "python_test" # 사용할 데이터베이스
)

if connection.is_connected():
    print("MySql에 성공적으로 연결되었습니다.")
else:
    print("연결 실패!")

connection.close()