import mysql.connector

# MYsql 연결 설정
connection = mysql.connector.connect(
    host = "localhost",      # MySql 서버주소를 넣을 예정
    user = "root",           # 사용자 이름
    password = "root",       # 비밀번호
    database = "python_test" # 사용할 데이터베이스
)

cursor = connection.cursor() # 데이터베이스 작업을 위한 객체 생성
# cursor = connection.cursor(dictionary=True) # 딕셔너리 형태로 가져오기

cursor.execute("SELECT * FROM users") # users 테이블 내 모든 데이터 조회

rows = cursor.fetchall() #조회 결과를 가져온다.
for row in rows:
    print(row) # 각 행 출력
    #print(rows['email'])



cursor.close()
connection.close() # finally 안에 넣으면 된다.