import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time
import sys
import re
import requests 
from dotenv import load_dotenv
from datetime import datetime
from openpyxl.styles import NamedStyle

load_dotenv()

class DiscordLogger:
    def __init__(self):
        self.token = os.getenv("DISCORD_TOKEN")
        self.channel_id = os.getenv("CHANNEL_ID")
        self.api_url = f"https://discord.com/api/v10/channels/{self.channel_id}/messages"
        self.headers = {
            "Authorization": f"Bot {self.token}",
            "Content-Type": "application/json"
        }

    def send_message(self, message, parse_mode=None):
        """傳送訊息到 Discord (Discord 預設支援 Markdown，不需特別指定 parse_mode)"""
        if not self.token or not self.channel_id:
            print("Discord Token 或 Channel ID 未設定！")
            return
            
        payload = {"content": message}
        try:
            response = requests.post(self.api_url, headers=self.headers, json=payload)
            response.raise_for_status() 
        except Exception as e:
            print(f"Discord 訊息發送失敗: {e}")

d_logger = DiscordLogger()
log_buffer = []

def log_print(message):
    """Print to console and append to log buffer."""
    print(message)
    log_buffer.append(str(message))

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('banklinker-473405-6be3b03228c7.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Bank").worksheet("總明細")

HEADLESS = False

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
        self.login_id = 0
        self.login_account = 0
        self.login_password = 0
        self.main_account = ""
        self.cash = 0
        self.exchange = 0
        self.stock = 0

Esun = Bank()
Esun.login_id = os.getenv("ESUN_ID")
Esun.login_account = os.getenv("ESUN_ACCOUNT")
Esun.login_password = os.getenv("ESUN_PASSWORD")

Cathay = Bank()
Cathay.login_id = os.getenv("CATHAY_ID")
Cathay.login_account = os.getenv("CATHAY_ACCOUNT")
Cathay.login_password = os.getenv("CATHAY_PASSWORD")

Line = Bank()
Line.login_id = os.getenv("LINE_ID")
Line.login_account = os.getenv("LINE_ACCOUNT")
Line.login_password = os.getenv("LINE_PASSWORD")

wait = WebDriverWait(driver, 20)

def EsunSpider():
    try:
        driver.get("https://ebank.esunbank.com.tw/index.jsp")

        driver.switch_to.default_content()
        wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe1"))
        )

        cust_input = wait.until(
            EC.visibility_of_element_located((By.ID, "loginform:custid"))
        )
        cust_input.clear()
        cust_input.send_keys(Esun.login_id)

        cust_input = wait.until(
            EC.visibility_of_element_located((By.ID, "loginform:name"))
        )
        cust_input.clear()
        cust_input.send_keys(Esun.login_account)

        cust_input = wait.until(
            EC.visibility_of_element_located((By.ID, "loginform:pxsswd"))
        )
        cust_input.clear()
        cust_input.send_keys(Esun.login_password)

        login_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "loginform:linkCommand"))
        )
        login_btn.click()

        span_el = wait.until(
            EC.presence_of_element_located((By.ID, "_0"))
        )
        Esun.main_account = span_el.text.strip()
        print(f"ESUNAccount：{Esun.main_account}")

        personal_balance_sheet = wait.until(
            EC.presence_of_element_located((By.XPATH, "//a[text()='個人資產負債表']"))
        )
        driver.execute_script("arguments[0].click();", personal_balance_sheet)

        balance_td = wait.until(
            EC.presence_of_element_located((By.ID, "fms01010a:twTd2"))
        )

        balance_text = balance_td.text.strip().replace(",", "")
        Esun.cash = int(balance_text)
        print(f"ESUNcash: {Esun.cash}")

        balance_td = wait.until(
            EC.presence_of_element_located((By.ID, "fms01010a:stockTd2"))
        )

        balance_text = balance_td.text.strip().replace(",", "")
        Esun.stock = int(balance_text)
        print(f"ESUNstock: {Esun.stock}")

        logout_button = driver.find_element(By.CSS_SELECTOR, "a.log_out")  
        logout_button.click()
    except Exception as e:
        log_print(f"Error in EsunSpider: {e}")
    
def CathaySpider():
    try:
        driver.get("https://www.cathaybk.com.tw/mybank/")

        cust_input = wait.until(
            EC.visibility_of_element_located((By.ID, "CustID"))
        )
        driver.execute_script("arguments[0].value = arguments[1];", cust_input, Cathay.login_id)

        cust_input = wait.until(
            EC.visibility_of_element_located((By.ID, "UserIdKeyin"))
        )
        driver.execute_script("arguments[0].value = arguments[1];", cust_input, Cathay.login_account)

        cust_input = wait.until(
            EC.visibility_of_element_located((By.ID, "PasswordKeyin"))
        )
        driver.execute_script("arguments[0].value = arguments[1];", cust_input, Cathay.login_password)

        loginButton = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and @class='btn no-print btn-fill js-login btn btn-fill w-100 u-pos-relative' and @onclick='NormalDataCheck()']"))
        )
        driver.execute_script("arguments[0].click();", loginButton)
        time.sleep(10)

        ###TWD
        button_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[data-evt="home_twd_overview"]'))
        )
        raw_text = button_element.text
        clean_text = raw_text.replace("TWD", "").replace(",", "").strip()
        Cathay.cash = int(clean_text)
        print(f"CATHAY_TWD: {Cathay.cash}")

        ###Foreign
        foreign_currency_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[data-evt="home_foreign_currency_overview"]'))
        )
        foreign_currency_text = foreign_currency_element.text
        clean_text = foreign_currency_text.replace("TWD", "").replace(",", "").strip().split()[0]
        Cathay.exchange = int(clean_text)
        print(f"CATHAYForeign: {Cathay.exchange}")

        ###STOCK
        xpath_selector = "//p[text()='投資']/parent::div/following-sibling::div[@class='css-iu1euh']/p"
        
        investment_element = wait.until(
            EC.visibility_of_element_located((By.XPATH, xpath_selector))
        )
        
        investment_text = investment_element.text
        clean_text = investment_text.replace("TWD", "").replace(",", "").strip().split()[0]
        Cathay.stock = int(clean_text)
        print(f"CATHAYstock: {Cathay.stock}")

        ###LOGOUT
        logout_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-evt="onlinebanking-logout"]'))
        )
        driver.execute_script("arguments[0].click();", logout_button)
    except Exception as e:
        log_print(f"Error in CathaySpider: {e}")


def LineSpider():
    try:
        driver.get("https://accessibility.linebank.com.tw/transaction")

        wait.until(EC.presence_of_element_located((By.ID, "nationalId"))).send_keys(Line.login_id)
        wait.until(EC.presence_of_element_located((By.ID, "userId"))).send_keys(Line.login_account)
        wait.until(EC.presence_of_element_located((By.ID, "pw"))).send_keys(Line.login_password)

        login_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='登入友善網路銀行']"))
        )
        login_btn.click()

        confirm_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='確定']"))
        )
        confirm_btn.click()

        dropdown_element = wait.until(
            EC.presence_of_element_located((By.ID, "account-dropdown"))
        )
        select = Select(dropdown_element)
        select.select_by_value("111003906466")


        p_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//p[contains(., '可用餘額')]"))
        )
        
        raw_text = p_element.text
        print(f"網頁上抓到的原始文字: '{raw_text}'")
        
        m = re.search(r"可用餘額\s*:\s*NT\$([0-9,]+)", raw_text)
        if m:
            Line.cash = int(m.group(1).replace(",", ""))
            print(f"LINEcash: {Line.cash}")
        else:
            print("警告：抓到了標籤，但格式無法解析！")
        
    except Exception as e:
        log_print(f"Error in LineSpider: {e}")
    
def JudgeColor(SheetRow, row):
    if SheetRow < 0:
        sheet.format(row, {'backgroundColor': {'red': 1, 'green': 0, 'blue': 0}})
    elif SheetRow > 0:
        sheet.format(row, {'backgroundColor': {'red': 0, 'green': 1, 'blue': 0}})
    elif SheetRow == 0:
        sheet.format(row, {'backgroundColor': {'red': 1, 'green': 1, 'blue': 1}})


try:
    log_print("Starting AutoAccount Task...")
    EsunSpider()
    CathaySpider()
    LineSpider()

    total_cash = Esun.cash + Cathay.cash + Line.cash
    total_exchange = Cathay.exchange
    total_stock = Esun.stock + Cathay.stock
    total_assets = total_cash + total_exchange + total_stock
    
    log_print(f"Total Cash: {total_cash}")
    log_print(f"Total Exchange: {total_exchange}")
    log_print(f"Total Stock: {total_stock}")
    log_print(f"Total Assets: {total_assets}")

    C3_value = int(sheet.cell(3, 3).value)
    D3_value = int(sheet.cell(3, 4).value) 
    E3_value = int(sheet.cell(3, 5).value) 
    F3_value = int(sheet.cell(3, 6).value) 

    cash_diff = total_cash - C3_value
    exchange_diff = total_exchange - D3_value
    stock_diff = total_stock - E3_value
    assets_diff = total_assets - F3_value

    current_date = datetime.now().strftime("%Y/%m/%d")
    current_time = datetime.now().strftime("%H:%M:%S")

    sheet.insert_row([current_date, current_time, 
                    total_cash, total_exchange, total_stock, total_assets, 
                    cash_diff, exchange_diff, stock_diff, assets_diff, " ", 
                    Esun.main_account, Esun.cash, Esun.exchange, Esun.stock, " ",  
                    Cathay.main_account, Cathay.cash, Cathay.exchange, Cathay.stock, " ",  
                    Line.main_account, Line.cash, Line.exchange, Line.stock
                    ], 3)

    G3_value = int(sheet.cell(3, 7).value)   
    H3_value = int(sheet.cell(3, 8).value)
    I3_value = int(sheet.cell(3, 9).value)
    J3_value = int(sheet.cell(3, 10).value)

    JudgeColor(G3_value, 'G3')
    JudgeColor(H3_value, 'H3')
    JudgeColor(I3_value, 'I3')
    JudgeColor(J3_value, 'J3')
    
    log_print("Task Completed Successfully.")
    
    def fmt(label, value=None, suffix=''):
        if value is None:
            return label
        if suffix == '%':
            return f"{label:<15}: {value:>13.2%}"
        return f"{label:<15}: {value:>13,}"

    summary_message = (
        f"```\n"
        f"AutoAccount Report \n"
        f"{current_date} {current_time}\n\n"
        f"{fmt('CASH RATIO', (total_cash + total_exchange )/total_assets,'%')}\n"
        f"-------------------------------\n" 
        f"{fmt('CHANGE')}\n"
        f"{fmt('Cash', cash_diff)}\n"
        f"{fmt('Exchange', exchange_diff)}\n"
        f"{fmt('Stock', stock_diff)}\n"
        f"{fmt('Assets', assets_diff)}\n"
        f"-------------------------------\n" 
        f"{fmt('SUM')}\n"
        f"{fmt('Cash', total_cash)}\n"
        f"{fmt('Exchange', total_exchange)}\n"
        f"{fmt('Stock', total_stock)}\n"
        f"{fmt('Assets', total_assets)}\n"
        f"-------------------------------\n"
        f"```" 
    )

    d_logger.send_message(summary_message)


except Exception as e:
    error_msg = f"Execution Failed: {e}"
    log_print(error_msg)
    d_logger.send_message(error_msg)
finally:
    try:
        driver.quit()
    except:
        pass
