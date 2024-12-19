import pandas as pd
import oracledb

# 1. 엑셀 파일 읽기
file_path = 'output_with_combined_round.xlsx'  # 엑셀 파일 경로
sheet_name = 'Sheet1'  # 엑셀 시트 이름
df = pd.read_excel(file_path, sheet_name=sheet_name)  # 데이터프레임으로 엑셀 파일 읽기
print(f"[INFO] 엑셀 파일 '{file_path}'의 '{sheet_name}' 시트를 성공적으로 불러왔습니다.")

# 2. Oracle 데이터베이스 연결 정보
db_user = 'dbp240107'
db_password = '81140'
db_host = 'dblab.dongduk.ac.kr'
db_port = 1521
db_service_name = 'orclpdb'

# 3. DSN 생성 및 DB 연결
try:
    dsn_tns = oracledb.makedsn(db_host, db_port, service_name=db_service_name)
    connection = oracledb.connect(user=db_user, password=db_password, dsn=dsn_tns)
    print("[INFO] Oracle DB에 성공적으로 연결되었습니다.")
except oracledb.DatabaseError as e:
    print(f"[ERROR] 데이터베이스 연결 실패: {e}")
    exit()

# 4. 커서 생성
cursor = connection.cursor()


# 데이터 삽입 쿼리 준비
insert_sql = """
    INSERT INTO examschedule (ROUND_ID, RECEPTION_START_DATE, RECEPTION_FINISH_DATE, RESULT_DATE, CERT_ID, CERT_NUM, EXAM_TYPE, EXAM_START_DATE, EXAM_END_DATE) 
    VALUES (:1, :2, :3, :4, :5, EXAMSCHEDULE_SEQ.NEXTVAL, :6, :7, :8)
"""

check_sql = """
    SELECT COUNT(*) 
    FROM EXAMSCHEDULE 
    WHERE ROUND_ID = :1 AND EXAM_TYPE = :2
"""
# 삽입할 데이터 리스트 준비
data_to_insert = []
for index, row in df.iterrows():
    data_to_insert.append((
        str(row['회차']).strip(),
        str(row['접수 시작']).strip(),
        str(row['접수 마감']).strip(),
        str(row['발표일자']).strip(),
        str(row['cert_id']).strip(),
        str(row['구분']).strip(),
        str(row['시험 시작 일자']).strip(),
        str(row['시험 마감일자']).strip(),
    ))

# 5. 데이터 삽입 실행
try:
    for data in data_to_insert:
        try:
            cursor.execute(insert_sql, data)  # 데이터 삽입
            connection.commit()  # 트랜잭션 커밋
        except oracledb.DatabaseError as e:
            error, = e.args
            print(f"[ERROR] 데이터 삽입 오류: {error.message}")
            print(f"[ERROR] 제약 조건 위반: {error.code}")  # 오류 코드와 메시지 출력
            connection.rollback()  # 롤백
            break

    print(f"[INFO] {len(data_to_insert)}개의 행이 데이터베이스에 삽입되었습니다.")
except oracledb.DatabaseError as e:
    print(f"[ERROR] 데이터 삽입 중 오류 발생: {e}")
    connection.rollback()  # 에러 발생 시 롤백

# 6. 연결 종료
cursor.close()  # 커서 종료
connection.close()  # 연결 종료
print("[INFO] 데이터베이스 연결이 종료되었습니다.")



