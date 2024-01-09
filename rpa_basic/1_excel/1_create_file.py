from openpyxl import Workbook
wb = Workbook() # 새 워크북 생성
ws = wb.active # 현재 활성화된 sheet 가져옴
ws.title = "NadoSheet" # sheet 의 이름을 변경
wb.save("sample.xlsx")
wb.close()

# https://www.youtube.com/watch?v=exgO1LFl9x8
print('저는 "나도코딩"입니다')
print("저는 \"나도코딩\"입니다")
print(" < 모두 성공 적으로 완성되었습니다 > ")