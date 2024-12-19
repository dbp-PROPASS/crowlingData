import pandas as pd
import json
# # # 파일 경로 (예시로 data.txt로 가정)
# txt_file_path = 'real.txt'  # TXT 파일 경로
# excel_file_path = 'output.xlsx'  # 저장될 엑셀 파일 경로

# # TXT 파일 읽기 (쉼표(,)로 구분된 데이터일 경우)
# df = pd.read_csv(txt_file_path, sep=',', header=None)

# # 열 이름 설정
# df.columns = ['자격증명', '기관', '회차', '접수 시작','접수 마감', '시험 시작 일자','시험 마감일자', '발표일자', '구분']

# # 날짜 형식 변경 함수 (2024.09.08 -> 2024/09/08)
# date_columns = ['접수 시작', '접수 마감', '시험 시작 일자', '시험 마감일자', '발표일자']

# # 날짜 형식 변환
# for column in date_columns:
#     df[column] = pd.to_datetime(df[column], errors='coerce').dt.strftime('%Y/%m/%d')

# # 엑셀로 저장
# df.to_excel(excel_file_path, index=False)

# print(f"TXT 파일이 엑셀로 저장되었습니다: {excel_file_path}")

##cert_id를 활용한 round_id 만들기 


# 엑셀 파일 경로
excel_file_path = 'output.xlsx'  # 엑셀 파일 경로
json_file_path = 'other_cert_ids.json'  # cert_id와 cert_name 매핑된 JSON 파일 경로
output_excel_path = 'output_with_cert_id.xlsx'  # 저장될 엑셀 파일 경로

# # 엑셀 파일 읽기
df = pd.read_excel(excel_file_path)

# JSON 파일 읽기 (cert_id와 cert_name 매핑된 파일)
with open(json_file_path, 'r', encoding='utf-8') as f:
    cert_mapping_data = json.load(f)

# JSON 데이터를 딕셔너리로 변환 (name을 키로, id를 값으로)
cert_mapping = {item['name']: item['id'] for item in cert_mapping_data}

# 자격증명에 따라 cert_id 추가
df['cert_id'] = df['자격증명'].map(cert_mapping)

# 엑셀로 저장 (cert_id 추가된 데이터)
df.to_excel(output_excel_path, index=False)

print(f"cert_id가 추가된 엑셀 파일이 저장되었습니다: {output_excel_path}")

# 엑셀 파일 경로
excel_file_path = 'output_with_cert_id.xlsx'  # 엑셀 파일 경로

# 엑셀 파일 읽기
df = pd.read_excel(excel_file_path)

# '회차' 필드에 cert_id와 회차 결합하여 넣기
df['회차'] = df['cert_id'].astype(str) + df['회차'].astype(str)

# 결과 엑셀 파일로 저장
output_file_path = 'output_with_combined_round.xlsx'
df.to_excel(output_file_path, index=False)

print(f"결과가 {output_file_path}로 저장되었습니다.")



