from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import datetime
import pyperclip
import pyautogui
import undetected_chromedriver as uc
import subprocess
import shutil
import os
import openpyxl
import warnings
import pandas as pd
import win32com.client


warnings.simplefilter("ignore")

subprocess.Popen(r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Users\sgn31\OneDrive\바탕 화면\No\Programming\selenium\디버그 모드"')
url = "https://kbland.kr/"
option = Options()
option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=option)
driver.maximize_window()
driver.get(url)

time.sleep(2)
aaa = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[7]/div/div[1]/button')
aaa.click()
time.sleep(2)
aaa = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[9]/div[2]/div/div/div[1]/div/section/div/div/div[3]/button')
aaa.click()
time.sleep(2)


# 구글로그인 클릭
aaa = driver.execute_script('document.querySelector("#app > div > div.memberArea.active > div.memberContent > div > div > div.scrollbar-inner.scroll-content > section > div.loginCont > div.btns > button.btn.btn-login.google > span").click();',aaa)
time.sleep(5)

# 지도 검색창으로 이동
aaa = driver.find_element(By.XPATH,'//*[@id="app"]/div/nav/div/button[2]')
aaa.click()
time.sleep(3)

# 검색창 누르기
aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
aaa.click()
time.sleep(3)

# 대상 주택 입력 및 검색
pyautogui.press('Tab',9)
# aaa = driver.find_element(By.XPATH,'//*[@id="__BVID__631"]/div[1]/h2')
# aaa.click()
# aaa = driver.execute_script('document.querySelector("#__BVID__575").click();',aaa)
aaa.send_keys('경기도 화성시 동탄순환대로 881')
pyautogui.press('enter')
time.sleep(3)

# 시세 Excel 다운로드
aaa = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[8]/div/div/span')
aaa.click()
time.sleep(3)
# aaa = driver.find_element(By.XPATH,'//*[@id="leftScroll"]/div/div[1]/div[2]/div/div[6]/h2')
# aaa.click()
pyperclip.copy('KB시세 다운로드') # 클립보드에 텍스트를 복사합니다. 
pyautogui.hotkey('ctrl', 'f')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
# pyautogui.press('Enter')
time.sleep(3)
aaa = driver.find_element(By.XPATH,'//*[@id="시세"]/div[2]/div/div[3]/button[2]')
aaa.click()
time.sleep(2)
aaa = driver.find_element(By.XPATH,'//*[@id="시세"]/div[2]/div/div[4]/div[2]/div/div/div[2]/button[1]/strong/em')
aaa.click()

time.sleep(3)

# 다운로드 파일 이름 변경 및 관리 폴더 쪽으로 위치 이동
d_today = datetime.date.today()

src = r"C:\Users\sgn31\Downloads"
dst = r"C:\Users\sgn31\OneDrive\바탕 화면\No\Programming\selenium\부동산 시세\동탄역푸르지오"

filename = "동탄역푸르지오 과거시세.xlsx"
new_filename = "(" + str(d_today) + ")" + filename

old_name = os.path.join(src, filename)
new_name = os.path.join(src, new_filename)

os.rename(old_name, new_name)

time.sleep(3)

shutil.move(src + "\\" + new_filename, dst + "\\" + new_filename)

new_path = dst + "\\" + new_filename
apart_path = r"C:\Users\sgn31\OneDrive\바탕 화면\No\Programming\selenium\부동산크롤링.xlsx"  # 부동산크롤링 엑셀 파일 위치

# wb = openpyxl.load_workbook(new_path)
# ws = wb.active

excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.open(new_path)
ws = wb.Activesheet

ws.Range("A1:D100").Copy()  # 시세자료 Data 복사

wc = excel.Workbooks.open(apart_path)   # 부동산크롤링 엑셀 실행
wd = wc.Activesheet

wd.Range("A1").Select()     
wd.Paste()                  # 부동산크롤링 엑셀파일 A2 선택 & 붙여넣기

for i in range(1,16):
    wd.Rows(1).EntireRow.Delete()

wd.Columns("A:D").Sort(Key1=wd.Range("A1"),Order1=1, Orientation=1)
wd.Rows(1).EntireRow.Insert()

# time.sleep(2)

wd.Range("A1").value = "동탄역푸르지오"
wd.Range("B1").value = "하위 평균가"
wd.Range("C1").value = "일반 평균가"
wd.Range("D1").value = "상위 평균가"

wb.save
wc.save

wb.close
wc.close

driver.close()

wf = openpyxl.load_workbook(apart_path)
wg = wf['Sheet1']

chart = openpyxl.chart.LineChart()
chart.title = "동탄역푸르지오 시세 추이"
chart.x_axis.title = "시세기준월"
chart.y_axis.title = "시세(만원)"
datas = openpyxl.chart.Reference(wg, min_col=2, min_row=1, max_col=4, max_row = 85)
chart.add_data(datas, from_rows=False, titles_from_data=True)
cats = openpyxl.chart.Reference(wg, min_col=1, min_row=2, max_col=1, max_row = 85)
chart.set_categories(cats)
chart.height = 15
chart.width = 30
wg.add_chart(chart, "F2")
wf.save(apart_path)
wf.close



# 다운로드 후 지정된 폴더까지 옮기기 완료!!

# # 이메일 입력
# pyautogui.typewrite('sgn3116@gmail.com')
# pyautogui.press('Enter')

# 다른 계정 사용
# aaa = driver.find_element(By.XPATH,'//*[@id="view_container"]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div/div/ul/li[1]/div/div[1]/div/div[2]/div[2]')
# aaa.click()

# # 기타 다른방법 로그인
# aaa = driver.find_element(By.CSS_SELECTOR,'.VfPpkd-vQzf8d')
# aaa.click()
# aaa = driver.find_element(By.XPATH,'//*[@id="view_container"]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div/ul/li[2]/div/div[1]/svg')
# aaa.click()

# # 비밀번호 입력
# pyautogui.typewrite('tjdwn6803')
# pyautogui.press('Enter')
# time.sleep(100)


# aaa = driver.find_element(By.XPATH,'//*[@id="view_container"]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div/div/ul/li[1]/div')
# aaa.click()
# time.sleep(2)
# aaa = driver.find_element(By.XPATH,'//*[@id="app"]/div/nav/div/button[2]')
# aaa.click()
# time.sleep(2)
# aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# aaa.click()

# # 주소 입력
# aaa.send_keys('경기도 화성시 동탄순환대로 881')
# aaa.send_keys(Keys.RETURN)

# aaa = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[8]/div/div/span')
# aaa.click()
# aaa = driver.find_element(By.XPATH,'//*[@id="complexSiseChart"]/div[1]/div[3]/button[2]')
# aaa.click()

# # 파일 다운로드
# aaa = driver.find_element(By.XPATH,'//*[@id="complexSiseChart"]/div[1]/div[4]/div[2]/div/div/div[2]/button[1]/strong/em')
# aaa.click()

# # 다운로드 excel 파일 실행

# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()
# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()
# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()
# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()
# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()
# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()
# # aaa = driver.find_element(By.XPATH,'//*[@id="searchArea"]/div/button[1]')
# # aaa.click()



# # # driver.execute_script("arguments[0].click();", aaa)

# # time.sleep(2)

