import mysql.connector

# MYsql 연결 설정
connection = mysql.connector.connect(
    host = "localhost",      # MySql 서버주소를 넣을 예정
    user = "root",           # 사용자 이름
    password = "root",       # 비밀번호
    database = "python_test" # 사용할 데이터베이스
)

cursor = connection.cursor() # 데이터베이스 작업을 위한 객체 생성

sql = "UPDATE users SET email = %s WHERE name = %s" # 이름을 받아 이메일을 업데이트
values = ("new_encore@example.com", "Encore")       # 업데이트할 데이터 값

cursor.execute(sql, values)   # 쿼리실행
connection.commit()           # 변경사항 저장

print(f'{cursor.rowcount}개의 행이 업데이트 되었습니다.') # 업데이트 된 행 수 출력

cursor.close()
connection.close()
