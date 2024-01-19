from openpyxl import Workbook
from openpyxl.drawing.image import Image
wb = Workbook()
ws = wb.active

# image 파일의 경로가 같은 폴더가 아니다 run 을 전용터미널에서 실행만 된다.
# 대화형 창에서는 실행이 안된다. 경로를 찾지 못 하는 에러가 발생함.
img = Image("rpa_basic/image.png")

# C3 위치에 img.png 파일의 이미지를 삽입
ws.add_image(img, "C3")

wb.save("sample_image.xlsx")
print(" < 모두 성공 적으로 완성되었습니다 > ")
