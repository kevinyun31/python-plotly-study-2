# pip install pyautogui <- 터미널에 입력하여 설치한다

import pyautogui

size = pyautogui.size()  # 현재 화면의 스크린 사이즈를 가져옴  
print(size)  # 가로, 세로 크기를 알 수 있음
# size[0] : width
# size[1] : height

print(" < 모두 성공 적으로 완성되었습니다 > ")
