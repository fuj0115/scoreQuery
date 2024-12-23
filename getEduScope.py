from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import cv2
import os
import numpy as np
import ddddocr

# 配置 Chrome 选项
chrome_options = Options()
#chrome_options.add_argument("--start-maximized")  # 最大化窗口

# 指定 ChromeDriver 路径
driver_path = r'D:\chromedriver-win64\chromedriver.exe'
service = Service(driver_path)

# 初始化 WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# 目标 URL
login_url = "https://www.eeagd.edu.cn/zkselfec/login/login.jsp"

try:
    # 打开登录页面
    start_time = time.time()
    driver.get(login_url)
    
    # 增加等待时间以确保页面完全加载
    time.sleep(0.5)  # 根据实际情况调整等待时间
    print(f"打开登录界面耗时:{time.time() - start_time:.4f} 秒")
    
    start_time = time.time()
    # 查找验证码图片元素
    captcha_element = driver.find_element(By.XPATH, '//*[@id="loginForm"]/div[3]/div/a/img')  # 根据实际情况修改定位方式
    
    # 获取验证码图片的位置和大小
    location = captcha_element.location
    size = captcha_element.size
    print(f"查找验证码位置耗时:{time.time() - start_time:.4f} 秒")
    
    start_time = time.time()
    # 截取整个屏幕
    screenshot = driver.get_screenshot_as_png()
    screenshot_np = np.frombuffer(screenshot, np.uint8)
    screenshot_cv = cv2.imdecode(screenshot_np, cv2.IMREAD_COLOR)
    print(f"截取整个屏幕耗时: {time.time() - start_time:.4f} 秒")
    
    # 计算验证码图片的实际位置
    start_time = time.time()
    left = int(location['x'])
    top = int(location['y'])
    right = int(left + size['width'])
    bottom = int(top + size['height'])
    print(f"计算验证码图片的实际位置耗时: {time.time() - start_time:.4f} 秒")
    
    # 截取验证码图片
    start_time = time.time()
    captcha_image = screenshot_cv[top:bottom, left:right]
    print(f"截取验证码图片耗时: {time.time() - start_time:.4f} 秒")
    
    # 保存原始验证码图片（可选）cl
    # 开始计时
    captcha_save_path = r"D:\getScore\captcha.png" #路径不能包含中文
    # 检查文件夹是否存在
    folder_path = os.path.dirname(captcha_save_path)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"文件夹 {folder_path} 不存在，已创建。")
    else:
        print(f"文件夹 {folder_path} 存在。")
    
    # 打印保存路径
    print(f"保存路径: {captcha_save_path}")
    start_time = time.time()
    cv2.imwrite(captcha_save_path, captcha_image)
    # 结束计时
    print(f"图片保存耗时: {time.time() - start_time:.4f} 秒")
    
     # 检查文件是否存在
    if not os.path.exists(captcha_save_path):
        raise FileNotFoundError(f"文件未成功保存: {captcha_save_path}")
    
    
    start_time = time.time()
    ocr = ddddocr.DdddOcr()  # 切换为第二套ocr模型为 ocr = ddddocr.DdddOcr(beta=True)
    print(f"初始化ddddocr耗时: {time.time() - start_time:.4f} 秒")
    
    start_time = time.time()
    with open(captcha_save_path, "rb") as file:
        image = file.read()
    captcha_result = ocr.classification(image)
    print(f"识别验证码耗时: {time.time() - start_time:.4f} 秒")
    print("验证码识别结果："+captcha_result)
    
     # 输入账号
    username_input = driver.find_element(By.ID, 'username')
    username_input.send_keys('030121102535')  # 替换为实际的用户名
    time.sleep(0.2)

    # 输入密码
    password_input = driver.find_element(By.ID, 'password')
    password_input.send_keys('Jiguanda0429.')  # 替换为实际的密码
    time.sleep(0.2)

    # 输入验证码
    captcha_input = driver.find_element(By.ID, 'addcode')
    captcha_input.clear()  # 清空验证码输入框（如果有默认值）
    captcha_input.send_keys(captcha_result)

    # 点击登录按钮
    login_button = driver.find_element(By.ID, 'logincommit')
    login_button.click()
    
    time.sleep(5)
    
    # 提交表单（假设按钮 ID 为 "submit-button"）
    # submit_button = driver.find_element(By.ID, "logincommit")  # 根据实际情况修改定位方式
    # submit_button.click()


except Exception as e:
    print(f"发生错误: {e}")

finally:
    # 关闭 WebDriver
    driver.quit()