from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

# 你的 Instagram 用户名
username = "sunshineshisan@outlook.com"
# 你的 Instagram 密码
password = "Aa123456.."
# 要关注的用户的 Instagram 个人主页链接
target_user_url = "https://www.instagram.com/target_username/"

# ChromeDriver 的路径，如果已经添加到环境变量，可省略该路径指定
chromedriver_path = "path/to/chromedriver"
service = Service(chromedriver_path)

# 创建 Chrome 浏览器实例
driver = webdriver.Chrome(service=service)

try:
    # 打开 Instagram 登录页面
    driver.get("https://www.instagram.com/accounts/login/")
    time.sleep(3)

    # 输入用户名和密码
    username_input = driver.find_element(By.NAME, "username")
    password_input = driver.find_element(By.NAME, "password")
    username_input.send_keys(username)
    password_input.send_keys(password)

    # 点击登录按钮
    login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
    login_button.click()
    time.sleep(5)

    # 打开要关注的用户的个人主页
    driver.get(target_user_url)
    time.sleep(3)

    # 点击关注按钮
    follow_button = driver.find_element(By.CSS_SELECTOR, "button[type='button']:contains('关注')")
    follow_button.click()
    print(f"已成功关注 {target_user_url} 的用户")

except Exception as e:
    print(f"发生错误: {e}")

finally:
    # 关闭浏览器
    time.sleep(3)
    driver.quit()