from selenium import webdriver
from selenium.webdriver.common.by import By
from html import unescape
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

chrome_options = Options()

user_data_dir = r'C:\Users\Neo\AppData\Local\Google\Chrome\User Data'
chrome_options.add_argument(f"user-data-dir={user_data_dir}")


driver = webdriver.Chrome(options=chrome_options)


driver.get("https://cokechat.azurewebsites.net/c/new")
time.sleep(5)

def get_element():

    text_box = driver.find_element(by=By.ID, value="prompt-textarea")
    send_button = driver.find_element(By.CSS_SELECTOR, "button[data-testid='send-button']")

    return(text_box,send_button)

try:

    get_element()

except:
    time.sleep(2)
    get_element()
# 定位元素
div_element = driver.find_element(By.CLASS_NAME, "markdown")

# 获取innerHTML属性的值
inner_html_content = div_element.get_attribute('innerHTML')

# 对HTML实体进行解码
decoded_html_content = unescape(inner_html_content)

print(decoded_html_content)

x = input('Pause..')
driver.quit()
