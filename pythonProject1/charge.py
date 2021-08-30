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
    print(left,right,top,bottom)
    path1 = '1.png'
    path2 = '2.png'
    driver.save_screenshot(path1)  # 截取当前窗口并保存图片
    im = Image.open(path1)  # 打开图片
    # 528 588 285 305 正常 //异常:   528 544 289 305 图片不显示//528 528 305 305  宽和高为0
    if left == right:
        return 'vcode error'
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
    driver.get(url)
    driver.find_element_by_id('ext-comp-1008').click()
    driver.find_element_by_xpath("//div[@class='x-combo-list-inner']/div[@class='x-combo-list-item'][1]").click()
    driver.find_element_by_id('terminalName').send_keys(com)
    driver.find_element_by_id('terminalPassword').send_keys(password)
    for i in range(0, 20):
        try:
            data = v_code()
            driver.find_element_by_id('randCode').clear()
            driver.find_element_by_id('randCode').send_keys(data)
            #异常
            if (repr(data)==r"'\x0c'"):
                driver.find_element_by_partial_link_text('换一个').click()
                continue
            #要注释掉,因为pytesseract 带有\n,如果是vcode error说明验证码图片异常,导致截图出错
            if(data == 'vcode error'):
                driver.find_element_by_class_name('x-btn-text').click()
            time.sleep(5)
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
def time_scope(start_t, end_t,post_url,cir):
    global header
    start_date = str(datetime.datetime.strptime(start_t, '%Y-%m-%d')).replace('-','')
    end_date = str(datetime.datetime.strptime(end_t, '%Y-%m-%d')).replace('-','')
    x = 0
    col = ['缴费金额', '账户余额', '缴费类型', '缴费时间', '缴费方式', '备注']
    df = pd.DataFrame(columns=col)
    form_data = {"startTime": start_date,"stopTime": end_date}
    response = requests.post(url=post_url, headers=header,data=form_data)
    if(response.json().get('exception')=="操作成功"):
        response = response.json().get('infos')
        for items in response:
            df.loc[x] = list(items.values())[0:6]
            x+=1
    else:
        print(response.json())
        print('数据量过多!')
        sys.exit(0)
    df['缴费金额']= df['缴费金额'].str.replace(',', '')
    df['账户余额'] = df['账户余额'].str.replace(',','')
    df['缴费类型'] = '缴费'
    df['缴费方式'] = '现金'
    df['运营商'] = cir
    df['插入时间'] = str(datetime.datetime.now())
    dtypedict = {
        '缴费金额': Float, '账户余额': Float , '缴费类型': NVARCHAR(30),
        '缴费时间': NVARCHAR(20), '缴费方式': NVARCHAR(10) , '备注': NVARCHAR(20),
        '运营商':NVARCHAR(10), '插入时间':NVARCHAR(30)
    }
    #插入方式
    df.to_sql('phone_charge2', engine, if_exists='append', index=False, dtype=dtypedict)
    print(i, '运营商有多少行', df.shape[0])

if __name__ == '__main__':
    conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
    engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
    driver = webdriver.Chrome()

    list2 = ['运营商', '网址', '账号', '密码','防护网','密码2','Referer','post_url']
    df2 = pd.DataFrame(columns=list2)
    # df2.loc[0] = ['星河', 'http://39.98.109.6:9090', '深圳臻信', 'yzWv3RAz','http://39.98.109.6:5101/', 'shenz523','http://39.98.109.6:9090/customer/chs/query-cus-payhistory.html','http://39.98.109.6:9090/customer/WebGetPayHistory']
    # df2.loc[1] = ['金枝', 'http://112.74.44.161:9090/customer/chs/index.html', '易联达', '888888','http://112.74.44.161:5101/','KEN421pzh$@!','http://112.74.44.161:9090/customer/chs/query-cus-payhistory.html','http://112.74.44.161:9090/customer/WebGetPayHistory']
    # df2.loc[2] = ['众信','http://47.107.103.58:9090/chs/','臻信催收61151398','888888','','','http://47.107.103.58:9090/chs/query-cus-payhistory.html','http://47.107.103.58:9090/chs/query-cus-payhistory.jsp']
    df2.loc[0] = ['小码', 'http://120.79.30.202:7348/customer/chs/index.html', '深圳市云诺信科技全通催收', 'mjjgKGsG','','','http://120.79.30.202:7348/customer/chs/query-cus-payhistory.html','http://120.79.30.202:7348/customer/WebGetPayHistory']
    # df2.loc[4] = ['京蓝宇', 'http://60.10.163.120:3459/chs/', '深圳融臻催收', 'JIsmVpkT','','','http://60.10.163.120:3459/chs/query-cus-payhistory.html','http://60.10.163.120:3459/chs/query-cus-payhistory.jsp']
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
        time_scope('2021-05-01','2021-08-28',i.post_url,i.运营商)


