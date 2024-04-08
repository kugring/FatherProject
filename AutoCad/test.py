import os
import re
from window_point import *
from cad_print import *
from plot_print import *
from cad_close import *

# 경로 설정
path = r'C:\Users\gram\Desktop\2.팔덕면'

# 해당 경로에 있는 폴더 목록 가져오기
folder_list = [name for name in os.listdir(path) if os.path.isdir(os.path.join(path, name))]

# 숫자로 시작하는 폴더와 그렇지 않은 폴더를 분리
numeric_folders = []
other_folders = []
for folder in folder_list:
    if re.match(r'^\d', folder):  # 숫자로 시작하는 경우
        numeric_folders.append(folder)
    else:
        other_folders.append(folder)

# 숫자로 시작하는 폴더를 숫자 기준으로 정렬
numeric_folders.sort(key=lambda x: int(re.match(r'^(\d+)', x).group()))

# 숫자로 시작하는 폴더와 그렇지 않은 폴더를 합침
sorted_folder_list = numeric_folders + other_folders

while True:
    # 사용자로부터 입력 받기
    while True:
        try:
            input_number = int(input(f"실행할 폴더 번호를 입력하세요 (1부터 {len(sorted_folder_list)}), 종료를 원하시면 0을 입력하세요: "))
            if 0 <= input_number <= len(sorted_folder_list):
                break
            else:
                print("입력한 번호가 유효하지 않습니다. 다시 시도하세요.")
        except ValueError:
            print("숫자를 입력하세요.")

    # 종료 조건 확인
    if input_number == 0:
        print("프로그램을 종료합니다.")
        break

    # 입력한 번호에 해당하는 폴더의 내용 확인
    folder_index = input_number - 1
    folder = sorted_folder_list[folder_index]
    folder_path = os.path.join(path, folder)
    dwg_files = [file for file in os.listdir(folder_path) if file.endswith('.dwg')]

    if dwg_files:
        print(f"{folder} 폴더의 .dwg 파일 목록:")
        for dwg_file in dwg_files:
            print(os.path.join(folder_path, dwg_file))

        # 첫 번째 .dwg 파일 실행
        dwg_file_path = os.path.join(folder_path, dwg_files[0])
        os.startfile(dwg_file_path)

        # 입력된 숫자에 따라서 함수 실행
        # 입력된 숫자에 따라서 함수 실행
        try:
            i = str(input("인쇄 개수: "))  # 입력된 숫자를 문자열로 변환하여 i에 대입
            cmd_input_click()
            zoom()

            # 추가 입력 받기
            while True:
                line_input = str(input("원도우 숫자를 누르시오 (종료를 원하시면 'q'를 입력하세요): "))
                if line_input.lower() == 'q':
                    cad_sub_close()  # 현재 열여있는 캐드 창을 3초동안  저장하고 닫는다.
                    break
                else:
                    cmd_input_click()
                    if line_input == '3':
                        line_off()
                    else:
                        line_on()
                    cmd_input_click()
                    plot()
                    plot_samsung_setting(i)
                    window(line_input)
                    print_check()
        except KeyboardInterrupt:
            print("입력이 취소되었습니다.")
    else:
        print(f"{folder} 폴더에 .dwg 파일이 존재하지 않습니다.")
