from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options


chrome_options = Options()

user_data_dir = r'C:\Users\Neo\AppData\Local\Google\Chrome\User Data'
chrome_options.add_argument(f"user-data-dir={user_data_dir}")


driver = webdriver.Chrome(options=chrome_options)


driver.get("https://cokechat.azurewebsites.net/c/new")
time.sleep(5)



text_box = driver.find_element(by=By.ID, value="prompt-textarea")
send_button = driver.find_element(By.CSS_SELECTOR, "button[data-testid='send-button']")

# 定位元素
    
text_box.send_keys('Hello')
time.sleep(1)

send_button.click()


x = input('Pause..')
driver.quit()
