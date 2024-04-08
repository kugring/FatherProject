from google.cloud import vision
from google.oauth2 import service_account
from pdf2image import convert_from_path
import os


def detect_text_from_pdf(pdf_directory):
    # 서비스 계정 키 파일의 경로
    credentials = service_account.Credentials.from_service_account_file(
        r'C:\Program Files\kugring-dadswork-8b30c7c36b89.json')

    # Vision API 클라이언트 생성
    client = vision.ImageAnnotatorClient(credentials=credentials)

    # 감지된 텍스트를 저장할 딕셔너리
    detected_texts = {'소유주': [], '지목': []}

    # PDF 파일만 필터링하여 순회
    for file in os.listdir(pdf_directory):
        if file.endswith(".pdf"):
            # PDF 파일 경로 설정
            pdf_path = os.path.join(pdf_directory, file)

            # PDF 파일을 이미지로 변환하여 임시 폴더에 저장합니다.
            images = convert_from_path(pdf_path)
            temp_folder = "temp_images"
            os.makedirs(temp_folder, exist_ok=True)
            image_path = f"{temp_folder}/temp_page.jpg"
            images[0].save(image_path, "JPEG")

            # 특정 좌표 범위와 해당 텍스트를 포함하는 리스트
            target_bounds = [
                ((1320, 604), (1635, 625)),
                ((145, 590), (210, 630)),
            ]

            # 이미지에서 텍스트를 감지합니다.
            with open(image_path, "rb") as image_file:
                content = image_file.read()
            image = vision.Image(content=content)
            response = client.text_detection(image=image)
            texts = response.text_annotations
            if texts:
                for text in texts:
                    description = text.description.strip()
                    if not description:
                        continue
                    vertices = [(vertex.x, vertex.y) for vertex in text.bounding_poly.vertices]
                    for (x1, y1), (x2, y2) in target_bounds:
                        if any(x1 <= x <= x2 and y1 <= y <= y2 for x, y in vertices):
                            if (x1, y1) == target_bounds[0][0] and (x2, y2) == target_bounds[0][1]:
                                detected_texts['소유주'].append(description)
                            elif (x1, y1) == target_bounds[1][0] and (x2, y2) == target_bounds[1][1]:
                                detected_texts['지목'].append(description)
                            break

            # 임시 이미지 파일을 삭제합니다.
            os.remove(image_path)
            os.rmdir(temp_folder)

            # 결과 출력
            owner_text = " ".join(detected_texts['소유주']).replace(" ", "")
            purpose_text = " ".join(detected_texts['지목']).replace(" ", "")
            print("소유주:", owner_text)
            print("지목:", purpose_text)

            # 감지된 텍스트 딕셔너리를 초기화합니다.
            detected_texts = {'소유주': [], '지목': []}

    # 감지된 텍스트 딕셔너리를 반환합니다.
    return detected_texts


# PDF 파일이 있는 디렉토리 경로
pdf_directory = r"C:\Users\gram\Desktop\7. 복흥면\4. 율평마을 농로 선형개선공사 1000\토지대장"
detect_text_from_pdf(pdf_directory)
