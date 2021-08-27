# driver.find_element_by_class_name('x-btn-text').click()
from selenium import webdriver
import datetime
import pandas as pd
import time
import selenium.common.exceptions as ex
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from PIL import  Image
import cv2
import pytesseract
import pymssql
from sqlalchemy import create_engine
from sqlalchemy.types import NVARCHAR, Float, Integer, Date, Numeric
import smtplib
from email.mime.text import MIMEText
import re
def v_code():
	full_img = r"./full_img.png"
	first_img = r"./img.png"
	driver.save_screenshot(full_img)
	image_site = driver.find_element_by_id('vfcode')
	left = image_site.location['x']
	top = image_site.location['y']
	right = image_site.location['x'] + image_site.size['width']
	bottom = image_site.location['y'] + image_site.size['height']
	image = Image.open(full_img)
	image = image.crop((left,top,right,bottom))
	image.save(first_img)
	img = Image.open(first_img).resize(((180,60))).convert("RGBA")
	pixdata = img.load()
	for y in range(img.size[1]):
		for x in range(img.size[0]):
			if pixdata[x, y][0] < 90:
				pixdata[x, y] = (0, 0, 0, 255)
	for y in range(img.size[1]):
		for x in range(img.size[0]):
			if pixdata[x, y][1] < 136:
				pixdata[x, y] = (0, 0, 0, 255)
	for y in range(img.size[1]):
		for x in range(img.size[0]):
			if pixdata[x, y][2] > 0:
				pixdata[x, y] = (255, 255, 255, 255)
	img.save(first_img)
	v_code = pytesseract.image_to_string(first_img)[:4]
	return v_code
def send_mail(titles,contexts):
    mail_host = 'smtp.163.com'
    mail_user = 'dq1601430804'
    mail_pass = 'TCNLAEMSSJGLUAOH'
    sender = 'dq1601430804@163.com'
    receivers = ['xvdanqi@szrongzhen.com']
    message = MIMEText(contexts,'plain','utf-8')
    message['Subject'] = titles
    message['From'] = sender
    message['To'] = receivers[0]
    # 授权码 ：TCNLAEMSSJGLUAOH
    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)
        smtpObj.login(mail_user, mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
        smtpObj.quit()
        print('send mail success.')
    except smtplib.SMTPException as e:
        print('send mail fail. error:', e)
def login_module(line,url,name,password):
    global driver,x
    #修改网址
    if name != '易联达':
        driver.get(url)
    #修改金枝的特殊登录
    driver.find_element_by_id('ext-comp-1008').click()
    driver.find_element_by_xpath("//div[@class='x-combo-list-inner']/div[@class='x-combo-list-item'][1]").click()
    #修改密码账号
    driver.find_element_by_id('terminalName').send_keys(name)
    driver.find_element_by_id('terminalPassword').send_keys(password)
    for i in range(0, 15):
        try:
            data = v_code()
            print(data)
            driver.find_element_by_id('randCode').send_keys(data)
            driver.find_element_by_class_name('x-btn-text').click()
            time.sleep(3)
            driver.find_element_by_id('menu-账户日消费').click()
        except ex.NoSuchElementException:
            print("验证码错了", i)
            time.sleep(2)
            driver.switch_to.default_content()
            driver.find_element_by_xpath("//div[@class='x-window x-window-plain x-window-dlg']//td[@class='x-btn-center']").click()
            driver.find_element_by_partial_link_text('换一个').click()
            continue
        else:
            break
    driver.switch_to.frame('tab账户日消费')
    for i in range(1,3):
        driver.find_elements_by_name('month')[i].click()
        driver.find_element_by_class_name('x-btn-wrap.x-btn').click()
        time.sleep(5)
        try:
            wait = WebDriverWait(driver, 60, 1)
            datas = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'x-grid3-row-table')),message="")
        except:
            # 修改中断
            continue
        if name != '易联达':
            for item in datas:
                list_all = item.text.splitlines()
                df.loc[x]=list_all[0:10]
                x+=1
        else:
            for item in datas:
                list_all = item.text.splitlines()
                df.loc[x, ['日期']] = list_all[0]
                df.loc[x,['费用总计']]=list_all[1]
                df.loc[x,['时长总计']]=list_all[2][0:8]
                df.loc[x, ['时长秒数']] = re.search('\(\d*\)',str(list_all[2])).group()
                df.loc[x, ['本地费用']] = list_all[3]
                df.loc[x, ['本地时长']] = list_all[4][0:8]
                df.loc[x, ['本地时长秒数']] = re.search('\(\d*\)',str(list_all[4][-3:])).group()
                df.loc[x, ['国内长途费用']] = list_all[5]
                df.loc[x, ['国内长途时长']] = list_all[6][0:8]
                df.loc[x, ['国内长途秒数']] = re.search('\(\d*\)',str(list_all[6][-7:])).group()
                x+=1
        time.sleep(5)
if __name__ == '__main__':
    time_start = datetime.datetime.now()
    conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
    engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
    circuit_list2 =['运营商','网址','账号','密码']
    circuit_df=pd.DataFrame(columns=circuit_list2)
    circuit_df.loc[0] = ['众信','http://47.107.103.58:9090/chs/','臻信催收61151398','888888']
    circuit_df.loc[1] = ['小码', 'http://120.79.30.202:7348/customer/chs/index.html', '深圳市云诺信科技全通催收', 'mjjgKGsG']
    circuit_df.loc[2] = ['金枝', 'http://112.74.44.161:9090/customer/chs/index.html', '易联达', '888888']
    circuit_df.loc[3] = ['京蓝宇', 'http://60.10.163.120:3459/chs/', '深圳融臻催收', 'JIsmVpkT']
    #记录线路爬取数量
    num_list = []
    # 循环遍历线路
    for i in circuit_df.itertuples():
        driver = webdriver.Chrome()
        if i.运营商 == '金枝' :
            driver.get('http://112.74.44.161:5101/')
            driver.find_element_by_id('passwd').send_keys('KEN421pzh$@!')
            current_window = driver.current_window_handle
            driver.find_element_by_id('button').click()
            time.sleep(5)
        col = ['日期', '费用总计', '时长总计', '时长秒数', '本地费用', '本地时长', '本地时长秒数', '国内长途费用', '国内长途时长', '国内长途秒数']
        df = pd.DataFrame(columns=col)
        #记录已爬取第几行
        x = 0
        login_module(i.运营商,i.网址, i.账号, i.密码)
        #修改时间范围
        # import datetime
        # date_start=str(datetime.datetime.now().date()-datetime.timedelta(days=2))
        # print(date_start)
        # time_scope()
        #修改运营商
        df['运营商'] = i.运营商
        df['插入时间'] = str(datetime.datetime.now())
        df['日期'] = df['日期'].str.replace('/','-')
        df['费用总计'] = df['费用总计'].str.replace(',', '')
        df['时长秒数']=re.search('.(\d*)',str(df['时长秒数'])).group(1)
        df['本地时长秒数']=re.search('.(\d*)',str(df['本地时长秒数'])).group(1)
        df['国内长途秒数'] = re.search('.(\d*)', str(df['国内长途秒数'])).group(1)
        print(df)
        dtypedict = {
            '日期':NVARCHAR(30),'费用总计':Float,'时长总计':NVARCHAR(30),
            '时长秒数':Integer,'本地费用':NVARCHAR(30),'本地时长':NVARCHAR(30),
            '本地时长秒数':Integer,'国内长途费用':NVARCHAR(30),'国内长途时长':NVARCHAR(30), '国内长途秒数':Integer,
            '运营商': NVARCHAR(10),'插入时间':NVARCHAR(30)
            }
        #修改插入方式
        df.to_sql('phone_dailybill2',engine,if_exists='append',index = False,dtype = dtypedict)
        print(i,'运营商有多少行',df.shape[0])


