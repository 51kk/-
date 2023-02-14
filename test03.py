import time


# test部分 -----------------------------------------------------------------------------------

from tkinter import *
import pandas as pd
import tkinter.filedialog
from selenium.webdriver.common.by import By
from webdriver_helper import get_webdriver
import os

import pywinauto

a = tkinter.filedialog.askopenfilename ()  # 让用户选择文件并且返回文件名
print (a)
nw = os.path.basename (a)
NW1 = str (nw)    #将文件名转为字符串
print (type (NW1))#测试类型
nw2 = NW1.replace ('.xlsx' or '.xls', '') #删除文件名后缀
print (nw2)

# wk= vb.load_workbook(r''+a)
file = pd.read_excel (r'' + a, converters={'bh': str})  #读取Excel文件  并且返回字符串类型
print (type (file))
print (file)
df1 = file.dropna ()  #删除空值"nan"
# print(file)
user_names = df1["序号"]  # 获取序号列数据保存在数组
print (type (user_names))
# print(user_names)
file_list = user_names.values.tolist ()
print (file_list)


# test部分 -----------------------------------------------------------------------------------


def excelFilesPath(path):
    '''
    path: 目录文件夹地址
    返回值：列表，pdf文件全路径
    '''
    filePaths = []  # 存储目录下的所有文件名，含路径
    for root, dirs, files in os.walk (path):
        for file in files:
            filePaths.append (os.path.join (root + "/", file))
    return filePaths


# 获取路径
# default_dir = r"文件路径"
# path = tkinter.filedialog.askdirectory(title=u'选择文件', initialdir=(os.path.expanduser((default_dir))))
path = r"D:\xml"

# k = PyKeyboard()

driver = get_webdriver ()

# 打开网站
driver.get ("http://cq.singlewindow.cn/Index.aspx")  #打开重庆单一窗口
time.sleep (12) #等待用户输入密码
# cookie_1 = {"Name": "ASP.NET_SessionId", "Value": "nv0ahli1dyrxt40i5dyn40zr"}

# driver.add_cookie(cookie_1)
# 定位注释
# time.sleep(5)
driver.get ("http://113.204.136.26:8180/cqsw/swProxy/deskserver/sw/deskIndex?menu_id=nexp")
driver.implicitly_wait (10)
# 定位子框架
driver.get ("https://www.singlewindow.cn/singlewindow/standard/app_new.jsp?area_id=500000")
el = driver.find_element (By.XPATH, "/html/body/div/div/div[2]/div[1]/ul/li[6]")
driver.implicitly_wait (10)
# # 定位物品通关
# el = driver.find_element(By.XPATH,"/html/body/div/div/div[2]/div[1]/ul/li[6]")
# el.click()
# time.sleep(3)
# # 定位快件通关
# el = driver.find_element(By.XPATH,"/html/body/div/div/div[2]/div[2]/div[6]/div[2]/ul/li[1]/a/div")
# el.click()
# time.sleep(5)


driver.get (
    "http://113.204.136.26:8180/cqsw/swProxy/swnexpwebserver/static/pages/nexp/NexpDec/nexpDecbillDeclare.html?ngBasePath=http%3A%2F%2F113.204.136.26%3A8180%2Fcqsw%2FswProxy%2Fswnexpwebserver%2F")

excellist = excelFilesPath (path)


Success = 0 #正常暂存
Successlist=[]

Failure = 0# 暂存失败
Failurelist=[]
#     test部分------------------------------------------------------------------------------
for lists in user_names: #遍历excel需要上传的序号
    print (lists)
    str_1 = int (lists)# 转为int类型
    print (str_1)
    str_2 = str (str_1) #转为字符串类型
    print (str_2)
    print (type (str_2))
    xmlname = nw2 + "-" + str_2 + "QP.xml"   #Excel 文件名 + 序号 + 后缀
    print (xmlname)
    el = driver.find_element (By.XPATH, "/html/body/div[2]/div[2]/button[4]") #定位上传文件
    print (el)
    el.click () #点击
    time.sleep (1)  #等待一秒
    app = pywinauto.Desktop ()  #获取窗口自动对象

    dialog = app['打开']  # 根据名字找到弹出窗口
    dialog["Toolbar3"].click ()
    dialog.type_keys (path) # 输入文件地址
    dialog.type_keys ('{ENTER}') #点击回车
    time.sleep (0.5)
    dialog["Edit"].type_keys (xmlname)  # 在输入框中输入值，excel文件名
    dialog["Button"].click ()
    time.sleep (0.5)

    try:  #尝试暂存
        # 定位并点击暂存
        el = driver.find_element (By.XPATH, "/html/body/div[2]/div[2]/button[2]")
        driver.implicitly_wait (20)
        el.click ()
        driver.implicitly_wait (20)

        # el = driver.find_element(By.NAME,"/html/body/div[3]/div[2]/div[2]/iframe[2]")
        # el.click()
        # time.sleep(3)

        # 点击是确认
        # alert = driver.switch_to.alert()
        # alert.accept()
        # iframe_el = driver.find_element(By.XPATH,r"//*[@id="layui-layer10"]/div[3]/a[1]")
        # driver.switch_to.iframe(0)
        driver.implicitly_wait (20)

        print (driver)
        el = driver.find_element (By.CLASS_NAME, "layui-layer-btn0").click ()
        print (el)
        driver.implicitly_wait (20)
        Success = Success+1
        Successlist.append(xmlname)

    except:#出错后
        el = driver.find_element (By.CLASS_NAME, "layui-layer-btn0").click ()
        print (el)
        driver.implicitly_wait (20)

        print(xmlname+"暂存错误")
        Failure = Failure+1
        Failurelist.append(xmlname)
print(len(Successlist))
print(len(Failurelist))
if len(Successlist) != 0 and len(Failurelist) != 0:
    Successlistcount =len(Successlist)   #计数成功的文件
    Successlist.append(Successlistcount)  #把长度加在最后一列
    Failurelistcount = len(Failurelist)  #计数失败的文件
    Failurelist.append(Failurelistcount)  #把长度加在最后一列

    Log = pd.DataFrame({"暂存失败": Failurelist},{"暂存成功": Successlist})

    logpath = "D:/xml/" + nw2 + ".xlsx"
    Log.to_excel(logpath)

    print("都有")
elif len(Failurelist)== 0 and len(Successlist) != 0:
    Successlistcount = len(Successlist)  # 计数成功的文件
    Successlist.append(Successlistcount)  # 把长度加在最后一列
    Log = pd.DataFrame({"暂存成功": Successlist})
    print("没有未暂存成功的")
elif len(Successlist)== 0 and len(Failurelist) != 0:
    Failurelistcount = len(Failurelist)  # 计数成功的文件
    Failurelist.append(Failurelistcount)  # 把长度加在最后一列
    Log = pd.DataFrame({"暂存成功": Failurelist})
    print("没有暂存成功的")
else:
    print("都没有")



Log = pd.DataFrame({"暂存失败": Failurelist})
logpath= "D:/xml/" + nw2 +".xlsx"
Log.to_excel(logpath)



    # else:
    #     el = driver.find_element (By.CLASS_NAME, "layui-layer-btn0").click ()
    #     print (el)
    #     driver.implicitly_wait (20)


#     test部分------------------------------------------------------------------------------
#
# for excelfis in excellist:
#
# #定位单票导入
#     print(excelfis)
#     el = driver.find_element(By.XPATH,"/html/body/div[2]/div[2]/button[4]")
#     print(el)
#     el.click()
#     time.sleep(2)
#     app = pywinauto.Desktop()
#
#     dialog = app['打开']  # 根据名字找到弹出窗口
#     dialog["Toolbar3"].click()
#     dialog.type_keys(path)
#     # k.type_string(r"D:\xml")
#
#     dialog.type_keys('{ENTER}')
#     # keyboard.press(key.enter)
#     # keyboard.press('enter')
#
#     time.sleep(1)
#
#     b = os.path.basename(excelfis)
#     dialog["Edit"].type_keys(b)  # 在输入框中输入值
#     dialog["Button"].click()
#     # 键盘录入文件名并确认
#     # app = application.Application()
#     # app.connect(class_name='#32770')
#     # app["Dialog"]["Edit1"].TypeKeys(r''+ excelfis)
#     # app["Dialog"]["Button1"].click()
#
#
#
#
#                                 # keyboard.write(r''+ excelfis)
#                                 # time.sleep(1)
#                                 # keyboard.press('enter')
#     time.sleep(1)
#     # 定位并点击暂存
#     el = driver.find_element(By.XPATH,"/html/body/div[2]/div[2]/button[2]")
#     time.sleep(1)
#     el.click()
#     time.sleep(2)
#
#
#     # el = driver.find_element(By.NAME,"/html/body/div[3]/div[2]/div[2]/iframe[2]")
#     # el.click()
#     # time.sleep(3)
#
#
#     # 点击是确认
#     # alert = driver.switch_to.alert()
#     # alert.accept()
#     # iframe_el = driver.find_element(By.XPATH,r"//*[@id="layui-layer10"]/div[3]/a[1]")
#     # driver.switch_to.iframe(0)
#     driver.implicitly_wait(20)
#     print(driver)
#     el= driver.find_element(By.CLASS_NAME,"layui-layer-btn0").click()
#     print(el)
#     time.sleep(1)
#
#











