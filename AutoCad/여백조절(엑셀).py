from openpyxl.worksheet.page import PageMargins
from openpyxl import load_workbook
from openpyxl.worksheet.page import PrintOptions

file_path = r"C:\Users\gram\Desktop\수량- 정동마을 위쪽 배수로 정비공사.xlsx"
sheet_name = "표지"  # 조절할 시트 이름


def center_margin(file_path, sheet_name):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]

    # 프린트 옵션 설정
    print_options = PrintOptions(horizontalCentered=True)
    ws.print_options = print_options

    # 여백 설정 (좁은 여백)
    margin = PageMargins(left=0.25197, right=0.25197, top=0.7512, bottom=0.7512, header=0.299, footer=0.299)
    ws.page_margins = margin

    wb.save(file_path)


center_margin(file_path, sheet_name)
