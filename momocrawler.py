import requests
import selenium
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By

import gspread
from urllib.parse import urljoin
from google.oauth2.service_account import Credentials

### 假設要查詢"DJI"
### 輸入搜尋名稱
search_name = input("請輸入搜尋名稱: ")

### 輸入價格範圍
min_price = input("請輸入最低價格: ")
max_price = input("請輸入最高價格: ")

### 建立價格搜尋區間的索引
price_index = f'{min_price}_{max_price}'

### 構建搜尋 URL
web = f'https://www.momoshop.com.tw/search/searchShop.jsp?keyword={search_name}&searchType=1&curPage=1&_isFuzzy=0&showType=chessboardType&priceIndex={price_index}'

### requests伺服器資料
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('User-Agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.70 Safari/537.36"')
driver = webdriver.Chrome(options=options)
driver.maximize_window()  ###最大化視窗
# driver.set_page_load_timeout(10)  #等10秒
driver.get(web)

# 填入價格範圍並點擊確認按鈕
price_input = driver.find_element(By.ID, 'priceS')
price_input.clear()
price_input.send_keys(min_price)

price_input = driver.find_element(By.ID, 'priceE')
price_input.clear()
price_input.send_keys(max_price)

confirm_btn = driver.find_element(By.CSS_SELECTOR, '.priceBtn')
confirm_btn.click()


### BeautifulSoup 正式開始
soup = BeautifulSoup(driver.page_source, "html.parser")

product = soup.find_all('a', class_='goodsUrl')  ### 找到目標區塊

#價格由低至高排序
product_sorted = sorted(product, key=lambda x: float(x.find('span', class_='price').text.replace('$', '').replace(',', '')))

name_list = []
slogan_list = []
price_list = [] 
url_list = []
link = []
for i in product_sorted:
    name_list.append(i.find('h3', class_="prdName").text)
    slogan_list.append(i.find('p', class_="sloganTitle").text)
    price_list.append(i.find('span', class_="price").text)
    url = urljoin(web, i['href'])  # 將相對路徑轉換為完整的 URL
    url_list.append(url)
    # link.append(i.find('a','href').text)#有問題,有要連結,和前5頁
    # url已修正

# 連線到 Google Sheets
scope = ['https://www.googleapis.com/auth/spreadsheets']
credentials = Credentials.from_service_account_file('./exalted-kit-391608-bdaf37439175.json', scopes=scope)
gc = gspread.authorize(credentials)

# 開啟指定的 Google Sheets 檔案
spreadsheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1oYLGZwnxuV2sf6HpQz84GX3awJkeb7b3MjbKh4OXaws/edit#gid=0')

# 選擇工作表
worksheet = spreadsheet.get_worksheet(0)

# 將資料寫入工作表
data = [['產品名稱', '副標題', '商品價格','連結']] + list(zip(name_list, slogan_list, price_list,url_list))
worksheet.clear()  # 清除工作表中原有的資料
worksheet.insert_rows(data, row=1)  # 將新資料插入到工作表中

driver.quit()
print("爬蟲結束")
