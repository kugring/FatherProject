from google.cloud import vision
from google.oauth2 import service_account
from pdf2image import convert_from_path
import os

# 서비스 계정 키 파일의 경로
credentials = service_account.Credentials.from_service_account_file(
    r'C:\Program Files\kugring-dadswork-8b30c7c36b89.json')

# Vision API 클라이언트 생성
client = vision.ImageAnnotatorClient(credentials=credentials)


def detect_text_from_pdf(pdf_path):
    # PDF 파일을 이미지로 변환하여 임시 폴더에 저장합니다.
    images = convert_from_path(pdf_path)
    temp_folder = "temp_images"
    os.makedirs(temp_folder, exist_ok=True)
    image_paths = []
    for i, image in enumerate(images):
        image_path = f"{temp_folder}/page_{i + 1}.jpg"
        image.save(image_path, "JPEG")
        image_paths.append(image_path)

    # 특정 좌표 범위와 해당 텍스트를 포함하는 리스트
    target_bounds = [
        ((1403, 566), (1878, 588))
    ]

    # 이미지에서 텍스트를 감지합니다.
    for image_path in image_paths:
        print(f"Text detected from {image_path}:")
        with open(image_path, "rb") as image_file:
            content = image_file.read()
        image = vision.Image(content=content)
        response = client.text_detection(image=image)
        texts = response.text_annotations
        if texts:
            for text in texts:
                vertices = [(vertex.x, vertex.y) for vertex in text.bounding_poly.vertices]
                for (x, y) in vertices:
                    if any(x1 <= x <= x2 and y1 <= y <= y2 for (x1, y1), (x2, y2) in target_bounds):
                        print(f'\n"{text.description}"')
                        print("bounds: {}".format(",".join([f"({x},{y})" for x, y in vertices])))
                        break
        else:
            print("No text detected.")

    # 임시 이미지 파일을 삭제합니다.
    for image_path in image_paths:
        os.remove(image_path)
    os.rmdir(temp_folder)


# PDF 파일 경로
pdf_path = r'C:\Users\gram\Desktop\354-1.pdf'

# 텍스트 감지 함수 호출
detect_text_from_pdf(pdf_path)
