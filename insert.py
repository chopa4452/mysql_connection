import mysql.connector

# MYsql 연결 설정
connection = mysql.connector.connect(
    host = "localhost",      # MySql 서버주소를 넣을 예정
    user = "root",           # 사용자 이름
    password = "root",       # 비밀번호
    database = "python_test" # 사용할 데이터베이스
)


cursor = connection.cursor() # 데이터베이스 작업을 위한 객체 생성

# 데이터 삽입 쿼리
sql = "INSERT INTO users (name, email) values (%s, %s)" # 사용자 데이터를 추가하는 SQL 쿼리
values = ("Encore", "encore@example.com")               # 값으로 사용할 데이터

cursor.execute(sql, values) # 쿼리 실행
connection.commit()         # 변경사항 커밋(저장한다)

print(f"{cursor.rowcount}개의 행이 추가되었습니다.") # 추가된 행 수 출력

cursor.close()
connection.close()