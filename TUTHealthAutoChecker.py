import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from time import sleep
from time import strftime
import os
import tkinter
from win10toast import ToastNotifier
import sys
import ctypes


# 清空文件夹
def removeall(path):
    filelist = os.listdir(path)
    for file in filelist:
        os.remove(path + '\\' + file)
    return


# 保存数据
def savepass():
    os.makedirs(imgpath)
    f = open(txtpath, 'w')
    f.write(fusername.get())
    f.write(', ')
    f.write(fpassword.get())
    f.write(', ')
    f.write(strftime('%Y%m%d'))
    f.close()
    root.destroy()


# 获取文件位置，只用于定位火狐驱动
def getPath(filename):
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    path = os.path.join(bundle_dir, filename)
    return path


imgpath = os.getcwd() + '\\TUTHealthAutoChecker\\images'
txtpath = os.getcwd() + '\\TUTHealthAutoChecker\\TUTHealthAutoChecker.txt'

if not os.path.exists(txtpath):

    # 创建windows定时任务
    if ctypes.windll.shell32.IsUserAnAdmin():
        # 已获得管理员权限，执行命令
        currentpath = os.getcwd() + '\\TUTHealthAutoChecker.exe'
        a = os.system(
            'SCHTASKS /Create /TN TUTHealthAutoChecker /TR ' + currentpath + ' /RU ADMINISTRATOR /SC ONLOGON /f')
    else:
        # 以管理员权限重启本程序
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit()

    # GUI
    newstart = True
    root = tkinter.Tk()
    fusername = tkinter.StringVar()
    fpassword = tkinter.StringVar()
    root.title('TUTHealthAutoChecker')
    root.geometry('550x300')
    tkinter.Label(text='请输入你的账号密码').grid(row=0, column=1)

    tkinter.Label(text='账号').grid(row=1, column=0)
    entry1 = tkinter.Entry(textvariable=fusername)
    entry1.grid(row=1, column=1)

    tkinter.Label(text='密码').grid(row=2, column=0)
    entry2 = tkinter.Entry(textvariable=fpassword)
    entry2.grid(row=2, column=1)
    tkinter.Button(text='确定', command=savepass).grid(row=3, column=1)

    tkinter.Label(text='').grid(row=6, column=1)
    tkinter.Label(text='请放心，你的账号密码不会被上传和泄露').grid(row=7, column=1)
    tkinter.Label(text='').grid(row=4, column=1)
    tkinter.Label(text='点击确定后，本程序会在开机时自动运行，但每天只能运行一次，而且不会再显示此页面\n\n'
                       '你的账号被保存在与本程序同目录下的TUTHealthAutoChecker\n'
                       '删除此文件夹即可重新设置账号密码，并再次运行').grid(row=5, column=1)
    tkinter.Label(text='本项目地址https://github.com/icynic/TUTHealthAutoChecker').grid(row=8, column=1)
    root.mainloop()
else:
    newstart = False

# 读取txt文件的数据
txt = open(txtpath)
strlist = txt.read().split(', ')
username = strlist[0]
password = strlist[1]
lasttime = strlist[2]
txt.close()

# 如果一天内第二次运行，直接终止程序
if lasttime == strftime('%Y%m%d') and newstart is False:
    sys.exit()

# 浏览器设定
option = webdriver.FirefoxOptions()
option.add_argument('--headless')
option.set_preference('browser.download.dir', imgpath)
option.set_preference('browser.download.folderList', 2)
driver = webdriver.Firefox(options=option, service_log_path=os.devnull,
                           executable_path=getPath('geckodriver.exe'))

# 打开网页
driver.get("http://wfw.tjut.edu.cn/orange/app/grmrjktzyd/index.html#/")

# 填用户名和密码
driver.implicitly_wait(120)
driver.find_element(By.NAME, "username").send_keys(username + Keys.ENTER
                                                   + password + Keys.ENTER)

# 我上报的
try:
    driver.implicitly_wait(45)
    driver.find_element(By.XPATH, '/html/body/div/div/div/div/div[2]/div/div/div/div/div[1]/div/div[3]/span') \
        .click()
except selenium.common.exceptions.NoSuchElementException:
    toaster = ToastNotifier()
    toaster.show_toast('TUTHealthAutoChecker',
                       '登录失败，可能是学校系统未响应，或账号密码错误，重设账号密码请删除文件夹TUTHealthAutoChecker',
                       duration=15,
                       icon_path='G:\\download\\ram.png')
    driver.quit()
    sys.exit()

# 最近记录
driver.implicitly_wait(10)
driver.find_element(By.XPATH,
                    '/html/body/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div[1]') \
    .click()

# 获取上次图片
driver.implicitly_wait(10)
url1 = driver.find_element(By.XPATH,
                           '/html/body/div/div/div/div/div[2]/div/div[1]/div/div/div/div/div/div[11]/div/div[2]/div/div/div/div/div/img') \
    .get_attribute('src')
url2 = driver.find_element(By.XPATH,
                           '/html/body/div/div/div/div/div[2]/div/div[1]/div/div/div/div/div/div[12]/div/div[2]/div/div/div/div/div/img') \
    .get_attribute('src')

# 清空图片
try:
    removeall(imgpath)
except FileNotFoundError:
    pass
driver.set_page_load_timeout(1)
driver.implicitly_wait(10)

# 下载图1
try:
    driver.get(url1)
except selenium.common.exceptions.TimeoutException:
    pass
driver.set_page_load_timeout(30000)

# 刷新
driver.refresh()

# 获取图1路径
image1 = os.listdir(imgpath)[0]

# 下载图2
driver.implicitly_wait(10)
driver.set_page_load_timeout(1)
try:
    driver.get(url2)
except selenium.common.exceptions.TimeoutException:
    pass
driver.set_page_load_timeout(3000)
driver.refresh()

# 获取图2路径
imglist = os.listdir(imgpath)
if image1 == imglist[0]:
    image2 = imglist[1]
else:
    image2 = imglist[0]

# 填写上报
driver.back()
driver.implicitly_wait(10)
driver.find_element(By.XPATH, '/html/body/div/div/div/div/div[2]/div/div/div/div/div[1]/div/div[1]') \
    .click()

# 复用上次
driver.implicitly_wait(10)
driver.find_element(By.XPATH,
                    '/html/body/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[1]/a') \
    .click()

# 上传图片
driver.implicitly_wait(10)
try:
    driver.find_element(By.XPATH,
                        '/html/body/div/div/div/div/div[2]/div/div[1]/div/div/div/div/div/div[9]/div/div[2]/div/div/div/div/input') \
        .send_keys(imgpath + '\\' + image1)
    driver.find_element(By.XPATH,
                        '/html/body/div/div/div/div/div[2]/div/div[1]/div/div/div/div/div/div[10]/div/div[2]/div/div/div/div/input') \
        .send_keys(imgpath + '\\' + image2)

    # 提交
    sleep(30)
    driver.find_element(By.XPATH,
                        '/html/body/div/div/div/div/div[2]/div/div[2]/div/div/div[1]/button').click()
    toaster = ToastNotifier()
    toaster.show_toast('TUTHealthAutoChecker', '打卡成功', duration=10, icon_path='G:\\download\\ram.png')

except selenium.common.exceptions.NoSuchElementException:
    toaster = ToastNotifier()
    toaster.show_toast('TUTHealthAutoChecker', '打卡失败，可能今日已提交或已逾期', duration=10, icon_path='G:\\download\\ram.png')

# 重新写入数据到txt文件
txt = open(txtpath, 'w')
txt.write(username)
txt.write(', ')
txt.write(password)
txt.write(', ')
txt.write(strftime('%Y%m%d'))
txt.close()

removeall(imgpath)
driver.quit()

'''
我修改了Python目录下 Lib\site-packages\selenium\webdriver\common\service.py 
内的start()函数，将 creationflags=self.creationflags 改为 creationflags=CREATE_NO_WINDOW
并添加了 from win32process import CREATE_NO_WINDOW
从而隐藏了浏览器驱动的dos窗口
'''
