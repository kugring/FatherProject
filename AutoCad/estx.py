from total_business_expense import *
from work_schedle import *
from title_content import *

'''
    *작동순서*
    
    1. 갑지_공사명_공사위치:     폴더경로에서 [수량상출서]엑셀을 열어서 [공사명,공사위치]를 작성하고 갑지(엑셀)을 해당경로에 저장해준다.
    2. 갑지_총사업비:           폴더경로의 [ESTX파일]을 열어서 [총공사비] 등등을 [갑지]에 입력해준다.
    3. 갑지_예정공정표:         폴더경로의 [ESTX파일]의 [내역서총괄표]에서 [품평, 합계]을 가져와 [갑지(엑셀)]로 입력해준다.
                              ㄴ(주의: 검은 그래프는 그려주지 않음)
                              
    {추가기능: 총사업비를 직접 전달하는 기능, 행정구역 정리해주는 코드
'''

# 폴더 경로
folder_path = r'C:\Users\gram\Desktop\7. 복흥면\4. 율평마을 농로 선형개선공사 1000'

갑지_공사명_공사위치(folder_path)
갑지_총사업비(folder_path)
갑지_예정공정표(folder_path)
print("끝!")
