import sys
import os


import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import re
import winreg
import zipfile
import requests
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

base_url = 'http://npm.taobao.org/mirrors/chromedriver/'
version_re = re.compile(r'^[1-9]\d*\.\d*.\d*')  # 匹配前3位版本号的正则表达式


def getChromeVersion():
    """通过注册表查询chrome版本"""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\Google\\Chrome\\BLBeacon')
        value, t = winreg.QueryValueEx(key, 'version')
        return version_re.findall(value)[0]  # 返回前3位版本号
    except WindowsError as e:
        # 没有安装chrome浏览器
        return "1.1.1"


def getChromeDriverVersion():
    """查询Chromedriver版本"""
    outstd2 = os.popen('chromedriver --version').read()
    try:
        version = outstd2.split(' ')[1]
        version = ".".join(version.split(".")[:-1])
        return version
    except Exception as e:
        return "0.0.0"


def getLatestChromeDriver(version):
    # 获取该chrome版本的最新driver版本号
    url = f"{base_url}LATEST_RELEASE_{version}"
    latest_version = requests.get(url).text
    print(f"与当前chrome匹配的最新chromedriver版本为: {latest_version}")
    # 下载chromedriver
    print("开始下载chromedriver...")
    download_url = f"{base_url}{latest_version}/chromedriver_win32.zip"
    file = requests.get(download_url)
    with open("chromedriver.zip", 'wb') as zip_file:  # 保存文件到脚本所在目录
        zip_file.write(file.content)
    print("下载完成.")
    # 解压
    f = zipfile.ZipFile("chromedriver.zip", 'r')
    for file in f.namelist():
        f.extract(file)
    print("解压完成.")


def checkChromeDriverUpdate():
    chrome_version = getChromeVersion()
    print(f'当前chrome版本: {chrome_version}')
    driver_version = getChromeDriverVersion()
    print(f'当前chromedriver版本: {driver_version}')
    if chrome_version == driver_version:
        print("版本兼容，无需更新.")
        return
    print("chromedriver版本与chrome浏览器不兼容，更新中>>>")
    try:
        getLatestChromeDriver(chrome_version)
        print("chromedriver更新成功!")
    except requests.exceptions.Timeout:
        print("chromedriver下载失败，请检查网络后重试！")
    except Exception as e:
        print(f"chromedriver未知原因更新失败: {e}")


checkChromeDriverUpdate()


# 添加当前文件路径到环境变量
curPath = os.path.abspath(os.path.dirname(__file__))
filePath = os.path.split(curPath)[0]
sys.path.append(filePath)
sys.path.extend([filePath + '\\' + i for i in os.listdir(filePath)])


def importfile():
    global Id, Pwd, nrows
    filename = 'config.xls'
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    nrows = table.nrows
    Id = table.col_values(colx=0)
    Pwd = table.col_values(colx=1)
    print("一共", nrows, "个账号", Id)
    print("一共", nrows, "个密码", Pwd)


def auto_click(n):
    n = nrows
    for i in range(n):
        id = Id[i]
        pwd = Pwd[i]
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--no-sandbox')
        options.add_argument('blink-settings=imagesEnabled=false')
        options.add_argument('--disable-gpu')
        s = Service('chromedriver.exe')
        driver = webdriver.Chrome(options=options, service=s)
        driver.get('http://dailyreport.hhu.edu.cn/pdc/form/list')
        time.sleep(2)
        stu_id = driver.find_element(By.ID, 'username')
        pwd_id = driver.find_element(By.ID, 'password')
        stu_id.send_keys(id)
        time.sleep(0.5)
        pwd_id.send_keys(pwd + Keys.RETURN)
        time.sleep(1)
        button1 = driver.find_element(By.PARTIAL_LINK_TEXT, '进入')
        button1.click()
        time.sleep(2)
        button2 = driver.find_element(By.ID, 'saveBtn')
        button2.click()
        time.sleep(2)
        print('完成第', i+1, '个')
        driver.close()


if __name__ == '__main__':
    importfile()
    auto_click(nrows)
    print('打卡成功')
    print('打卡成功')
    print('打卡成功')
    print('Written by Nicolarjibardo')
    time.sleep(3)
    sys.exit()
