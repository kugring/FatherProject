import os


# 경로 설정
path = r'C:\Users\gram\Desktop\7. 복흥면'

# 해당 경로에 있는 폴더 목록 가져오기
folder_list = [name for name in os.listdir(path) if os.path.isdir(os.path.join(path, name))]

while True:
    # 사용자로부터 입력 받기
    while True:
        try:
            input_number = int(input(f"실행할 폴더 번호를 입력하세요 (1부터 {len(folder_list)}), 종료를 원하시면 0을 입력하세요: "))
            if 0 <= input_number <= len(folder_list):
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
    folder = folder_list[folder_index]
    folder_path = os.path.join(path, folder)
    dwg_files = [file for file in os.listdir(folder_path) if file.endswith('.dwg')]

    if dwg_files:
        print(f"{folder} 폴더의 .dwg 파일 목록:")
        for dwg_file in dwg_files:
            print(os.path.join(folder_path, dwg_file))

        # 첫 번째 .dwg 파일 실행
        dwg_file_path = os.path.join(folder_path, dwg_files[0])
        os.startfile(dwg_file_path)
    else:
        print(f"{folder} 폴더에 .dwg 파일이 존재하지 않습니다.")

