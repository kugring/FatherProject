import xlrd

# 엑셀 파일 열기
workbook = xlrd.open_workbook(r'C:\Users\gram\Desktop\7. 복흥면\4. 율평마을 농로 선형개선공사 1000\율평마을 농로 선형개선공사.xls')

# 원가계산서 시트 선택
sheet = workbook.sheet_by_index(0)  # 첫 번째 시트 선택 (원가계산서가 처음에 있어서 가능)

# 탐색할 단어들
target_words = ["총     공    사     비", "관   급   자   재   대", "도        급        액", "부   가   가   치   세",
                "총        원        가"]

# 엑셀 시트를 순회하면서 각 단어를 찾음
for target_word in target_words:
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            cell_value = sheet.cell_value(row_index, col_index)
            if cell_value == target_word:

                # 발견된 셀의 오른쪽에 있는 값을 가져옴
                if col_index + 1 < sheet.ncols:  # 열 범위를 넘어가지 않도록 체크
                    next_cell_value = sheet.cell_value(row_index, col_index + 2)
                    print(f"{target_word.replace(' ','')}: {next_cell_value}")

# 엑셀 파일 닫기
workbook.release_resources()
