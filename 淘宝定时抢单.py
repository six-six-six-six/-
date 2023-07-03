import datetime
import time

import pyperclip
from selenium import webdriver
import win32com.client
import pyautogui
import xlrd

def mouseClick(clickTimes, lOrR, img, reTry):

    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            # location = pyautogui.locateCenterOnScreen(img)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("img: " + img + " 未找到匹配图片,0.1秒后重试")
            time.sleep(0.01)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.01)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.01)

def pay(img):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("单击左键", img)
        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("双击左键", img)
        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("右键", img)
            # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('1', '5', '7', '3', '0', '5', interval=0.02)
            print("输入:", inputValue)
            # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
        i += 1

def driver():
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    # 设置Chrome和Chrome驱动路径
    option = webdriver.ChromeOptions()
    option.binary_location = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
    option.add_argument('headless')
    # open_browser = webdriver.Chrome("C:\\Program Files\\Google\\Chrome\\Application\\driver\\chromedriver.exe",
    #                                 options=option)
    open_browser = webdriver.Chrome("C:\\Program Files\\Google\\Chrome\\Application\\driver\\chromedriver.exe")

    open_browser.get("https://www.taobao.com")
    time.sleep(2)
    open_browser.find_element_by_link_text("亲，请登录").click()
    time.sleep(2)
    open_browser.get("https://cart.taobao.com/cart.htm")
    time.sleep(2)
    return open_browser


def main():
    """
    *****淘宝定时抢单*****
    ！！！一切的前提！！！->网速OK
    tips:
    kill_time为抢单的时间、购物车为所需购买的物品（设定为选中购物车所有商品）
    需要人工执行的操作：
    1.淘宝二维码登录（随便什么时候登录就行，登录后就开始计算时间抢单了）
    2.提交订单后需打开支付宝二维码扫描支付
    """
    open_browser = driver()
    time.sleep(5)
    while True:
        try:
            # 选中后购物车所有物品
            if open_browser.find_element_by_id("J_SelectAll1"):
                time.sleep(5)
                open_browser.find_element_by_id("J_SelectAll1").click()
                print("已选中物品")
                break
        except:
            print("没有找到全选按钮")
    time.sleep(2)

    while True:
        now_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        print(now_time)
        if now_time > kill_time:
            break
    # ---结算---
    while True:
        try:
            if open_browser.find_element_by_link_text("结 算"):
                open_browser.find_element_by_link_text("结 算").click()
                print("结算成功, 时间为: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')))
                break
        except:
            print("尚未结算")
        time.sleep(0.01)
    # ---提交订单---
    while True:
        try:
            if open_browser.find_element_by_link_text('提交订单'):
                open_browser.find_element_by_link_text('提交订单').click()
                print("提交订单成功，请尽快付款, 时间为: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')))
                break
        except:
            print("已结算，请及时提交订单")
        time.sleep(0.01)
    # ---支付---
    # 数据检查
    pay(sheet1)
    print("付款成功, 付款时间为: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')))



if __name__ == '__main__':
    file = 'cmd.xls'
    kill_time = "2023-06-19 22:00:00.000000"
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    # 数据检查
    main()
    # driver()
