import os
from openpyxl import load_workbook

# 엑셀 파일 경로
file_path = r"C:\Users\gram\Desktop\8. 재혁면\@복흥 24면용역\정동마을 안길 아스콘 포장공사 외27건 실시설계용역.xlsm"

# 파일이 위치한 상위 상위 폴더 경로 가져오기
parent_parent_folder = os.path.abspath(os.path.join(file_path, os.pardir, os.pardir))  # 상위 상위 폴더의 경로

# 엑셀 파일 열기
workbook = load_workbook(file_path)
sheet = workbook.active

# 시작 셀과 끝 셀 설정
start_cell = "B4"
end_cell = "J4"

# 시작 셀과 끝 셀 좌표 구하기
start_column = ord(start_cell[0]) - 64  # 엑셀 열의 알파벳을 숫자로 변환
start_row = int(start_cell[1:])
end_column = ord(end_cell[0]) - 64
end_row = int(end_cell[1:])

# 찾을 단어 리스트
target_words = ['연번', '사 업 명(부기)', '사업비']

# 시작 셀부터 끝 셀까지 순회하면서 target_words를 포함하는 열 찾기
target_columns = []
for col in range(start_column, end_column + 1):
    cell_value = sheet.cell(row=start_row, column=col).value
    if cell_value in target_words and cell_value != '사업개요':  # '사업개요'는 제외
        target_columns.append(col)

# 폴더 생성
for row in sheet.iter_rows(min_row=7, min_col=min(target_columns), max_col=max(target_columns)):
    formatted_data = ''
    for idx, cell in enumerate(row):
        if idx + start_column in target_columns:
            if cell.value is None or cell.value == '':  # 빈 문자열인 경우 처리
                value = ''
            elif isinstance(cell.value, int) and cell.value >= 10000:  # 10000 이상인 경우에만 처리
                str_value = str(cell.value)
                if str_value.endswith('0'):
                    value = int(str_value[:-4])
                else:
                    value = cell.value
            else:
                value = cell.value
            # 연번과 사업 명 등 데이터 추가
            if idx + start_column == target_columns[0]:
                formatted_data += f'{value}.'
            elif idx + start_column == target_columns[1]:
                formatted_data += f' {value}'
            elif idx + start_column == target_columns[2]:
                formatted_data += f' {value}'

    # 폴더명 생성
    folder_name = formatted_data.strip()  # 폴더명에서 앞뒤 공백 제거
    folder_path = os.path.join(parent_parent_folder, folder_name)
    os.makedirs(folder_path, exist_ok=True)  # 폴더가 없으면 생성하고, 있으면 pass

print("폴더 생성이 완료되었습니다.")
