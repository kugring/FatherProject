import win32com.client as win32

공사명 = "관평마을 주차장"
공사장소 = "순창군 복흥면 정산리 582번지 일원"
행정구역 = "복흥면"
배수공 = "벤치플륨관설치 0.4*0.4  L=24.0m"


# HWP 파일 경로
file_path = r"C:\Users\gram\Desktop\한글TEST.hwp"

# 한/글 객체 생성
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
# 보안을 위한 코드
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

# 한/글 창을 표시합니다.
hwp.XHwpWindows.Item(0).Visible = True

# 문서 열기
hwp.Open(file_path)

hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option = hwp.HParameterSet.HFindReplace
option.FindString = "(공사이름)"
option.ReplaceString = 공사명
option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option = hwp.HParameterSet.HFindReplace
option.FindString = "(공사장소)"
option.ReplaceString = 공사장소
option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option = hwp.HParameterSet.HFindReplace
option.FindString = "(행정구역)"
option.ReplaceString = "      ".join(행정구역)
option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option = hwp.HParameterSet.HFindReplace
option.FindString = "(배수공)"
option.ReplaceString = 배수공
option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

# 수정된 내용을 저장 (나중에 경로는 타켓폴더로 지정해야한다.)
hwp.SaveAs(r"C:\Users\gram\Desktop\편집된_설계서5.hwp")

# 작업이 끝났으면 한/글을 닫습니다.
hwp.Quit()

print("편집이 완료되었습니다.")
