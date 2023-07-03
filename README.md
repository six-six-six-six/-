# 提前准备：
pyperclip 1.8.2  
selenium 3.141.0  
pyautogui 0.9.53  
xlrd 2.0.1  
opencv 4.6.0  
kill_time是设定的抢购时间  
修改dirver()函数中option.binary_location变量-谷歌浏览器的地址，open_browser变量-谷歌浏览器对应版本的驱动地址  
修改pay()函数中pyautogui.hotkey的密码  
修改main()函数的kill_time-开始抢购时间  



# 执行：
1、直接执行 淘宝定时抢单.py  
2、打开淘宝页面后需要手动扫码登录  
3、进入购物车页面选中物品后等待kill_time  
4、超过kill_time则开始抢购：结算→提交订单→支付  

# 注意事项：
1、若卡在无法匹配3中png，需要手动运行一遍流程，获取适应于本机的png  
2、需要提前运行代码，选中物品后才开始等待kill_time  
3、脚本会选中购物栏中所有物品  
4、请在保证网速1Mb/s以上运行，否则无法及时反应

