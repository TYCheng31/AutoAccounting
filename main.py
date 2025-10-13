import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time
import sys
import re
from dotenv import load_dotenv
from datetime import datetime
from openpyxl.styles import NamedStyle

load_dotenv()

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('banklinker-473405-6be3b03228c7.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Bank").sheet1  

HEADLESS = True

chrome_options = Options()
if HEADLESS:
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")
else:
    chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(options=chrome_options)

class Bank:
    def __init__(self):
        self.LoginId = 0
        self.LoginAccount = 0
        self.LoginPassword = 0
        self.MainAccount = 0
        self.Cash = 0
        self.Exchange = 0
        self.Stock = 0

Esun = Bank()
Esun.LoginId = os.getenv("ESUN_ID")
Esun.LoginAccount = os.getenv("ESUN_ACCOUNT")
Esun.LoginPassword = os.getenv("ESUN_PASSWORD")

Cathay = Bank()
Cathay.LoginId = os.getenv("CATHAY_ID")
Cathay.LoginAccount = os.getenv("CATHAY_ACCOUNT")
Cathay.LoginPassword = os.getenv("CATHAY_PASSWORD")

Line = Bank()
Line.LoginId = os.getenv("LINE_ID")
Line.LoginAccount = os.getenv("LINE_ACCOUNT")
Line.LoginPassword = os.getenv("LINE_PASSWORD")


def EsunSpider():
    driver.get("https://ebank.esunbank.com.tw/index.jsp")
    wait = WebDriverWait(driver, 20)

    driver.switch_to.default_content()
    WebDriverWait(driver, 20).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe1"))
    )

    cust_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "loginform:custid"))
    )
    cust_input.clear()
    cust_input.send_keys(Esun.LoginId)

    cust_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "loginform:name"))
    )
    cust_input.clear()
    cust_input.send_keys(Esun.LoginAccount)

    cust_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "loginform:pxsswd"))
    )
    cust_input.clear()
    cust_input.send_keys(Esun.LoginPassword)

    login_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "loginform:linkCommand"))
    )
    login_btn.click()

    span_el = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "_0"))
    )
    Esun.MainAccount = span_el.text.strip()
    print("ESUNAccount：", Esun.MainAccount)

    personal_balance_sheet = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//a[text()='個人資產負債表']"))
    )
    driver.execute_script("arguments[0].click();", personal_balance_sheet)

    balance_td = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "fms01010a:twTd2"))
    )

    balance_text = balance_td.text.strip().replace(",", "")
    Esun.Cash = int(balance_text)
    print("ESUNCash:", Esun.Cash)

    balance_td = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "fms01010a:stockTd2"))
    )

    balance_text = balance_td.text.strip().replace(",", "")
    Esun.Stock = int(balance_text)
    print("ESUNStock:", Esun.Stock)

    logout_button = driver.find_element(By.CSS_SELECTOR, "a.log_out")  # 使用CSS選擇器定位
    logout_button.click()
def CathaySpider():
    driver.get("https://www.cathaybk.com.tw/mybank/")
    wait = WebDriverWait(driver, 120)

    cust_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "CustID"))
    )
    driver.execute_script("arguments[0].value = arguments[1];", cust_input, Cathay.LoginId)

    cust_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "UserIdKeyin"))
    )
    driver.execute_script("arguments[0].value = arguments[1];", cust_input, Cathay.LoginAccount)

    cust_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "PasswordKeyin"))
    )
    driver.execute_script("arguments[0].value = arguments[1];", cust_input, Cathay.LoginPassword)

    loginButton = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and @class='btn no-print btn-fill js-login btn btn-fill w-100 u-pos-relative' and @onclick='NormalDataCheck()']"))
    )
    driver.execute_script("arguments[0].click();", loginButton)

    link_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, "//a[contains(@onclick, 'AutoGoMenu') and @class='link u-fs-14']"))
    )

    Cathay.MainAccount = link_element.text
    print("CATHAYAccount:", Cathay.MainAccount)

    balance_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "TD-balance"))
    )
    balance_text = balance_element.text
    Cathay.Cash = int(balance_text.replace(",", ""))  # 先去掉逗號，再轉換為整數
    print("CATHAYCash:", Cathay.Cash)

    tabFTD = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "tabFTD"))
    )
    driver.execute_script("arguments[0].click();", tabFTD)

    balance_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "FTD-balance"))
    )
    balance_text = balance_element.text
    Cathay.Exchange = int(balance_text.replace(",", ""))  # 先去掉逗號，再轉換為整數
    print("CATHAYExchange:", Cathay.Exchange)

    tabFUND = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "tabFUND"))
    )
    driver.execute_script("arguments[0].click();", tabFUND)

    fund_balance_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.ID, "FUND-balance"))
    )
    fund_balance_text = fund_balance_element.text

    # 移除逗號並轉換為數字
    Cathay.Stock = int(fund_balance_text.replace(",", ""))  # 先去掉逗號，再轉換為整數
    print("CATHAYStock:", Cathay.Stock)

    logout_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@onclick='IsNeedCheckReconcil()']"))
    )
    driver.execute_script("arguments[0].click();", logout_button)
def LineSpider():
    driver.get("https://accessibility.linebank.com.tw/transaction")
    wait = WebDriverWait(driver, 20)

    wait.until(EC.presence_of_element_located((By.ID, "nationalId"))).send_keys(Line.LoginId)
    wait.until(EC.presence_of_element_located((By.ID, "userId"))).send_keys(Line.LoginAccount)
    wait.until(EC.presence_of_element_located((By.ID, "pw"))).send_keys(Line.LoginPassword)

    login_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@title='登入友善網路銀行']"))
    )
    login_btn.click()

    confirm_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@title='確定']"))
    )
    confirm_btn.click()

    wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(., '可用餘額')]")))

    h2 = driver.find_element(By.XPATH, "//h2[contains(., '主帳戶')]")
    txt = re.sub(r"\s+", "", h2.text)                    
    Line.MainAccount = re.search(r"[（(]([0-9\-]+)[)）]", txt).group(1)
    print("LINEAccount:", Line.MainAccount)

    p = driver.find_element(By.XPATH, "//p[contains(., '可用餘額')]")
    ptxt = re.sub(r"\s+", "", p.text)                    
    m = re.search(r"NT\$?([0-9,]+(?:\.[0-9]+)?)", ptxt)
    available_display = f"NT${m.group(1)}"
    Line.Cash = int(m.group(1).replace(",", ""))
    print("LINECash:", Line.Cash)
def JudgeColor(SheetRow, row):
    if SheetRow < 0:
        sheet.format(row, {'backgroundColor': {'red': 1, 'green': 0, 'blue': 0}})
    elif SheetRow > 0:
        sheet.format(row, {'backgroundColor': {'red': 0, 'green': 1, 'blue': 0}})



EsunSpider()
CathaySpider()
LineSpider()

TotalCash = Esun.Cash + Cathay.Cash + Line.Cash
TotalExchange = Cathay.Exchange
TotalStock = Esun.Stock + Cathay.Stock
TotalAssets = TotalCash + TotalExchange + TotalStock
print("TotalCash:", TotalCash)
print("TotalExchange:", TotalExchange)
print("TotalStock:", TotalStock)
print("TotalAssets:", TotalAssets)

D6_value = int(sheet.cell(6, 4).value)   
E6_value = int(sheet.cell(6, 5).value)
F6_value = int(sheet.cell(6, 6).value)
G6_value = int(sheet.cell(6, 7).value)

I6_value = TotalCash - D6_value
J6_value = TotalExchange - E6_value
K6_value = TotalStock - F6_value
L6_value = TotalAssets - G6_value

current_date = datetime.now().strftime("%Y/%m/%d")
current_time = datetime.now().strftime("%H:%M:%S")
current_date2 = datetime.now().strftime("%m/%d")

sheet.insert_row([current_date], 2) 
sheet.insert_row([current_time, '玉山銀行', Esun.MainAccount, Esun.Cash, Esun.Exchange, Esun.Stock], 3) 
sheet.insert_row([" ", '國泰銀行', Cathay.MainAccount, Cathay.Cash, Cathay.Exchange, Cathay.Stock], 4) 
sheet.insert_row([" ", '連線銀行', Line.MainAccount, Line.Cash, Line.Exchange, Line.Stock], 5) 
sheet.insert_row([" ", " ", " ", TotalCash, TotalExchange, TotalStock, TotalAssets, current_date2, I6_value, J6_value, K6_value, L6_value], 6)


I6_value = int(sheet.cell(6, 9).value)   
J6_value = int(sheet.cell(6, 10).value)
K6_value = int(sheet.cell(6, 11).value)
L6_value = int(sheet.cell(6, 12).value)

JudgeColor(I6_value, 'I6')
JudgeColor(J6_value, 'J6')
JudgeColor(K6_value, 'K6')
JudgeColor(L6_value, 'L6')
