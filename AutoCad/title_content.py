from openpyxl import load_workbook
import os

'''

    *사용방법*
     - 엑셀(수량)파일에서 ['공사명','공사위치']를 추출해서
       엑셀(갑지)파일의 ['사업개요']에 저장한다.

    *주의사항*
    - 수량산출서 파일명을 "수량- "으로 정확히 앞에 기입해줘야한다.
'''

folder_path = r"C:\Users\gram\Desktop\7. 복흥면\4. 율평마을 농로 선형개선공사 1000"


def 갑지엑셀_생성(folder_path):
    # 엑셀 파일에서 데이터 가져오기
    def get_data_from_excel(file_path, sheet_name, cell):
        wb = load_workbook(filename=file_path)
        ws = wb[sheet_name]
        data = ws[cell].value
        wb.close()
        return data

    # 엑셀 파일에 데이터 입력하기
    def set_data_to_excel(file_path, sheet_name, cell, data):
        wb = load_workbook(filename=file_path)
        ws = wb[sheet_name]
        ws[cell] = data
        wb.save(filename=file_path)
        wb.close()

    # 폴더 내의 모든 파일을 확인하여 "수량 -"이라는 단어가 들어가는 엑셀 파일을 찾음
    found_file = None
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') and "수량-" in file:
            quantity_file_path = os.path.join(folder_path, file)
            # "수량- "을 제거한 파일명 가져오기
            공사명, _ = os.path.splitext(file.replace("수량- ", ""))
            break
    else:

        print("폴더 내에 '수량 -'이라는 단어가 들어가는 엑셀 파일이 없습니다.")
        exit()

    if found_file:
        print(f"찾은 파일명: {공사명}")

    title_excel_path = os.path.join(folder_path, '2.갑지- 복흑면23-4.xlsx')

    # 수량 파일에서 데이터 가져오기
    공사위치 = get_data_from_excel(quantity_file_path, '용지조서', 'B2')

    # 갑지 파일에 공사위치 데이터 입력하기
    set_data_to_excel(title_excel_path, '사업개요', 'C2', 공사위치)

    # 갑지 파일에 공사명 데이터 입력하기
    set_data_to_excel(title_excel_path, '사업개요', 'C1', 공사명)

    print("데이터 입력이 완료되었습니다.")
