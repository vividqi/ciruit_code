from selenium import webdriver
import datetime
import pandas as pd
import time
from PIL import Image
import pytesseract
import pymssql
from sqlalchemy import create_engine
import requests
from sqlalchemy.types import NVARCHAR, Float, Integer, Date, Numeric
import sys
import cv2
def v_code():
    img = driver.find_element_by_id('vfcode')
    left = int(img.location['x'])  # 获取图片左上角坐标x
    top = int(img.location['y'])  # 获取图片左上角y
    right = int(img.location['x'] + img.size['width'])  # 获取图片右下角x
    bottom = int(img.location['y'] + img.size['height'])
    path1 = '1.png'
    path2 = '2.png'
    driver.save_screenshot(path1)  # 截取当前窗口并保存图片
    im = Image.open(path1)  # 打开图片
    im = im.crop((left, top, right, bottom))  # 截图验证码
    im.save(path2)
    img = cv2.imread(path2, 0)
    thresh = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    data = pytesseract.image_to_string(thresh, config='--psm 6 digits')
    return data
def sec_to_clock(sec):
    m, s = divmod(int(sec), 60)
    h, m = divmod(m, 60)
    strr = "%02d:%02d:%02d" % (h, m, s)
    return strr
def login(url,com,password):
    time.sleep(0.5)
    driver.get(url)
    driver.find_element_by_id('ext-comp-1008').click()
    driver.find_element_by_xpath("//div[@class='x-combo-list-inner']/div[@class='x-combo-list-item'][1]").click()
    driver.find_element_by_id('terminalName').send_keys(com)
    driver.find_element_by_id('terminalPassword').send_keys(password)
    for i in range(0, 15):
        try:
            data = v_code()
            driver.find_element_by_id('randCode').clear()
            driver.find_element_by_id('randCode').send_keys(data)
            #要注释掉,因为有\n
            # driver.find_element_by_class_name('x-btn-text').click()
            time.sleep(2)
            driver.find_element_by_id('menu-账户话单查询').click()
        except:
            print(i,":验证码错了")
            time.sleep(2)
            driver.switch_to.default_content()
            driver.find_element_by_xpath("//div[@class='x-window x-window-plain x-window-dlg']//td[@class='x-btn-center']").click()
            driver.find_element_by_partial_link_text('换一个').click()
            continue
        else:
            break
    # 获取cookie
    cookie = driver.get_cookies()[0].get('name') + "=" + driver.get_cookies()[0].get('value')
    return cookie
def time_scope(start_t, days,post_url,cir):
    global header
    start_date = datetime.datetime.strptime(start_t, '%Y-%m-%d')
    for j in range(0, days):
        to_date = start_date + datetime.timedelta(days=j)
        # 每天置零
        x = 0
        col = ['主叫号码', '被叫号码', '起始时间', '通话时长', '通话费用', '套餐时长', '套餐费用', '通话类型', '计费方式']
        df = pd.DataFrame(columns=col)
        for i in range(9, 23, 1):
            to_start = str(to_date + datetime.timedelta(hours=i)).replace('-','')
            print(to_start)
            to_end = str(to_date + datetime.timedelta(hours=i, minutes=59, seconds=59)).replace('-','')
            form_data = {"startTime": to_start,"stopTime": to_end,"callerE164": "","calleeE164": ""}
            response = requests.post(url=post_url, headers=header,data=form_data)
            if(response.json().get('exception')=="操作成功"):
                response = response.json().get('infos')
                for items in response:
                    df.loc[x] = items.values()
                    x+=1
            else:
                for k in range(0,6):
                    to_start = str(to_date + datetime.timedelta(hours=i, minutes=k*10)).replace('-', '')
                    print(to_start)
                    to_end = str(to_date + datetime.timedelta(hours=i, minutes=k*10+9, seconds=59)).replace('-', '')
                    print(to_end)
                    form_data = {"startTime": to_start, "stopTime": to_end, "callerE164": "", "calleeE164": ""}
                    response = requests.post(url=post_url, headers=header, data=form_data)
                    if(response.json().get('exception')=="操作成功"):
                        response = response.json().get('infos')
                        for items in response:
                            df.loc[x] = items.values()
                            x += 1
                    else:
                        print('数据量过多!')
                        sys.exit(0)
        df['通话时长']= df['通话时长'].apply(lambda x: sec_to_clock(x))
        df['通话类型'] = '国内长途'
        df['计费方式'] = '主叫'
        df['运营商'] = cir
        df['插入时间'] = str(datetime.datetime.now())
        dtypedict = {
            '主叫号码': NVARCHAR(30), '被叫号码': NVARCHAR(30), '起始时间': NVARCHAR(30),
            '通话时长': NVARCHAR(20), '通话费用': Float, '套餐时长': NVARCHAR(20),
            '套餐费用': Float, '通话类型': NVARCHAR(20), '计费方式': NVARCHAR(20),
            '运营商': NVARCHAR(20), '插入时间': NVARCHAR(30)
        }
        #插入方式
        df.to_sql('phone_record', engine, if_exists='append', index=False, dtype=dtypedict)
        print(cir,i, '运营商有多少行', df.shape[0])
if __name__ == '__main__':
    conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
    engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
    driver = webdriver.Chrome()

    list2 = ['运营商', '网址', '账号', '密码','防护网','密码2','Referer','post_url']
    df2 = pd.DataFrame(columns=list2)
    df2.loc[0] = ['星河', 'http://39.98.109.6:9090', '深圳臻信', 'yzWv3RAz','http://39.98.109.6:5101/', 'shenz523','http://39.98.109.6:9090/customer/chs/query-cus-cdr.html','http://39.98.109.6:9090/customer/WebGetCustomerCdr']
    df2.loc[1] = ['金枝', 'http://112.74.44.161:9090/customer/chs/index.html', '易联达', '888888','http://112.74.44.161:5101/','KEN421pzh$@!','http://112.74.44.161:9090/customer/chs/query-cus-cdr.html','http://112.74.44.161:9090/customer/WebGetCustomerCdr']
    df2.loc[2] = ['众信','http://47.107.103.58:9090/chs/','臻信催收61151398','888888','','','http://47.107.103.58:9090/chs/query-cus-cdr.html','http://47.107.103.58:9090/chs/query-cus-cdr.jsp']
    df2.loc[3] = ['小码', 'http://120.79.30.202:7348/customer/chs/index.html', '深圳市云诺信科技全通催收', 'mjjgKGsG', '', '', 'http://120.79.30.202:7348/customer/chs/query-cus-cdr.html','http://120.79.30.202:7348/customer/WebGetCustomerCdr']
    df2.loc[4] = ['京蓝宇', 'http://60.10.163.120:3459/chs/', '深圳融臻催收', 'JIsmVpkT','','','http://60.10.163.120:3459/chs/query-cus-cdr.html','http://60.10.163.120:3459/chs/query-cus-cdr.jsp']

    for i in df2.itertuples():
        if i.运营商 == '星河' or i.运营商 == '金枝':
            driver.get(i.防护网)
            driver.find_element_by_id('passwd').send_keys(i.密码2)
            current_window = driver.current_window_handle
            driver.find_element_by_id('button').click()
            time.sleep(5)
        cookie=login(i.网址, i.账号, i.密码)
        header = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
            "Cookie": cookie,
            "Referer": i.Referer
        }
        time_scope('2021-08-26',3,i.post_url,i.运营商)



