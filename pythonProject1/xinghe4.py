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
            driver.find_element_by_class_name('x-btn-text').click()
            time.sleep(2)
            driver.find_element_by_id('menu-??????????????????').click()
        except:
            print(i,":???????????????")
            time.sleep(2)
            driver.switch_to.default_content()
            driver.find_element_by_xpath(
                "//div[@class='x-window x-window-plain x-window-dlg']//td[@class='x-btn-center']").click()
            driver.find_element_by_partial_link_text('?????????').click()
            continue
        else:
            break
    # ??????cookie
    cookie = driver.get_cookies()[0].get('name') + "=" + driver.get_cookies()[0].get('value')
    return cookie
def time_scope(start_t, days,post_url,cir):
    global header
    start_date = datetime.datetime.strptime(start_t, '%Y-%m-%d')
    for j in range(0, days):
        to_date = start_date + datetime.timedelta(days=j)
        # ????????????
        x = 0
        col = ['????????????', '????????????', '????????????', '????????????', '????????????', '????????????', '????????????', '????????????', '????????????']
        df = pd.DataFrame(columns=col)
        for i in range(9, 23, 1):
            to_start = str(to_date + datetime.timedelta(hours=i)).replace('-','')
            print(to_start)
            to_end = str(to_date + datetime.timedelta(hours=i, minutes=59, seconds=59)).replace('-','')
            form_data = {"startTime": to_start,"stopTime": to_end,"callerE164": "","calleeE164": ""}
            response = requests.post(url=post_url, headers=header,data=form_data)
            if(response.json().get('exception')=="????????????"):
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
                    if(response.json().get('exception')=="????????????"):
                        response = response.json().get('infos')
                        for items in response:
                            df.loc[x] = items.values()
                            x += 1
                    else:
                        print('???????????????!')
                        sys.exit(0)
        df['????????????']= df['????????????'].apply(lambda x: sec_to_clock(x))
        df['????????????'] = '????????????'
        df['????????????'] = '??????'
        df['?????????'] = cir
        df['????????????'] = str(datetime.datetime.now())
        dtypedict = {
            '????????????': NVARCHAR(30), '????????????': NVARCHAR(30), '????????????': NVARCHAR(30),
            '????????????': NVARCHAR(20), '????????????': Float, '????????????': NVARCHAR(20),
            '????????????': Float, '????????????': NVARCHAR(20), '????????????': NVARCHAR(20),
            '?????????': NVARCHAR(20), '????????????': NVARCHAR(30)
        }
        #????????????
        df.to_sql('phone_record', engine, if_exists='append', index=False, dtype=dtypedict)
        print(i, '?????????????????????', df.shape[0])
if __name__ == '__main__':
    conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
    engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
    driver = webdriver.Chrome()

    list2 = ['?????????', '??????', '??????', '??????','?????????','??????2','Referer','post_url']
    df2 = pd.DataFrame(columns=list2)
    df2.loc[0] = ['??????', 'http://39.98.109.6:9090', '????????????', 'yzWv3RAz','http://39.98.109.6:5101/', 'shenz523','http://39.98.109.6:9090/customer/chs/query-cus-cdr.html','http://39.98.109.6:9090/customer/WebGetCustomerCdr']
    for i in df2.itertuples():
        driver.get(i.?????????)
        driver.find_element_by_id('passwd').send_keys(i.??????2)
        current_window = driver.current_window_handle
        driver.find_element_by_id('button').click()
        time.sleep(5)
        cookie=login(i.??????, i.??????, i.??????)
        header = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
            "Cookie": cookie,
            "Referer": i.Referer
        }
        time_scope('2021-08-18',8,i.post_url,i.?????????)


