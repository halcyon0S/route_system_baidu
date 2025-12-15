# test_browser.py
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
import time

edge_options = Options()
edge_options.add_argument('--no-sandbox')
edge_options.add_argument('--disable-dev-shm-usage')
edge_options.add_argument('--remote-debugging-port=55555')

print("创建浏览器...")
driver = webdriver.Edge(service=Service(), options=edge_options)

print("访问百度...")
driver.get("https://www.baidu.com")

print("等待 5 秒...")
time.sleep(5)

print("检查 URL:", driver.current_url)
print("测试成功！按 Enter 关闭...")
input()

driver.quit()