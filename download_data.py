from selenium import webdriver
from selenium.webdriver.support.ui import Select
import os
import time
import datetime

county = "新竹市 (HsinchuCity)"
start_year = input("起始年分: ")
start_month = int(input("起始月份: "))   # required 01, 02, ...,12 instead of 1, 2, 3, ..., 12
start_date = int(input("起始日期: "))    # required 1, 2, ..., 30 instead of 01, 02, ..., 30
end_year = input("結束年分: ")
end_month = input("結束月份: ")
end_date = input("結束日期: ")

#### 字串處理：網站索引是用string資料型態去比對 ####
start_date = str(start_date)
if((start_month-10) < 0):
   start_month = "0" + str(start_month)  # 1->01,  2->02
else:
   start_month = str(start_month)        # 11->11, 12->12
#################################################


#### datetime套件，都是用integer資料型態作為函式的參數 ####
strt = datetime.date(int(start_year), int(start_month), int(start_date))
end = datetime.date(int(end_year), int(end_month), int(end_date))
########################################################
time_interval = (end - strt)  # datatime套件自己有定義時間加減法(包含閏年等)
counter = int(time_interval.days) 



#### datatime自己是一種資料型態(class)，和string一起print的時候必續強制轉換成string資料型態 ####
print("Downlaoding data from " + str(strt) + " to " + str(end))
print("Total: %d files" %(counter+1))
############################################################################################



# chromePath = "C:\\Users\\Heyward\\Desktop\\下載氣象資料\\chromedriver.exe"
chromePath = os.getcwd() + "\\chromedriver.exe"
print(chromePath)
driver = webdriver.Chrome(chromePath)
url = "https://e-service.cwb.gov.tw/HistoryDataQuery/index.jsp"
driver.get(url)
time.sleep(5)

### 選取城市&測站 ###
county_btn = Select(driver.find_element_by_id("stationCounty"))
station_btn = Select(driver.find_element_by_id("station"))
county_btn.select_by_visible_text(county)
station_btn.select_by_visible_text("新竹市東區 (Dongqu Hsinshu City)")


### 選取年月:下拉選單是動態網頁，要模擬點擊time_table ###
driver.find_element_by_class_name("ui-datepicker-trigger").click()

year_btn = Select(driver.find_element_by_class_name("ui-datepicker-year"))
year_btn.select_by_visible_text(start_year)
time.sleep(2) #選完年分後網頁又會重新載入，故等待載入再獲取month_btn

month_btn = Select(driver.find_element_by_class_name("ui-datepicker-month"))
month_btn.select_by_visible_text(start_month)
time.sleep(1)

### 選取日期 ###
date_btn_list = driver.find_elements_by_class_name("ui-state-default")
for date_btn in date_btn_list:
   if(date_btn.text == start_date):
      date_btn.click()

### 點擊查詢紐，切換至新分頁 ###
driver.find_element_by_id("doquery").click()
driver.switch_to_window(driver.window_handles[1])
driver.maximize_window()
### 點擊CSV下載紐，再點擊下一天 ###

while(counter>=0):
   print("%s 下載中...剩餘 %d 筆資料" %(str(strt), counter))
   driver.find_element_by_id("downloadCSV").click()
   driver.find_element_by_id("nexItem").click()
   counter = counter -1
   strt = strt + datetime.timedelta(days=1)
   time.sleep(1)
   
print("下載完成!")
driver.close()
driver.switch_to_window(driver.window_handles[0])
driver.close()