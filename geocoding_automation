from openpyxl import load_workbook
from geopy.geocoders import Nominatim
import time

# 엑셀 파일 열기
file_path = "과제.xlsx"
wb = load_workbook(file_path)
sheet = wb.active

# T열과 U열 사이에 두 개의 열 삽입
sheet.insert_cols(idx=21, amount=2)  

# Nominatim 객체 생성
geolocator = Nominatim(user_agent="geoapiExercise")

# T열에 있는 대학명을 읽고 위도, 경도를 U열과 V열에 입력
for row in range(2, sheet.max_row + 1):  # 2부터 시작하여 헤더를 건너뛰기
    university_name = sheet.cell(row=row, column=20).value  
    if university_name and university_name.strip():  # None과 공백 문자열 처리
        try:
            location = geolocator.geocode(university_name.strip())
            if location:
                sheet.cell(row=row, column=21).value = location.latitude  
                sheet.cell(row=row, column=22).value = location.longitude 
            else:
                sheet.cell(row=row, column=21).value = "N/A"
                sheet.cell(row=row, column=22).value = "N/A"
        except Exception as e:
            print(f"Error for {university_name}: {e}")
            sheet.cell(row=row, column=21).value = "Error"
            sheet.cell(row=row, column=22).value = "Error"
        time.sleep(0.2)  # API 호출 간 딜레이 추가

# 파일 저장
wb.save(file_path)
