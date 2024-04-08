import numpy as np
from pdf2image import convert_from_path
import pytesseract
import cv2

# Tesseract 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# PDF 파일에서 이미지 생성
pages = convert_from_path(r'C:\Users\gram\Desktop\되나.pdf', poppler_path=r'C:\Program Files\poppler-24.02.0\Library\bin')

# 이미지를 이용하여 텍스트 추출
text = ""
for page in pages:
    # 이미지를 NumPy 배열로 변환
    page_np = np.array(page)

    # 이미지 전처리
    gray = cv2.cvtColor(page_np, cv2.COLOR_RGB2GRAY)
    _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

    # 텍스트 추출
    text += pytesseract.image_to_string(binary, lang='kor')

# 텍스트 출력
print(text)
