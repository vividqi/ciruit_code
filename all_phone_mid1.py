from selenium import webdriver
import datetime
import pandas as pd
import time
from PIL import Image
import pytesseract
import pymssql
from sqlalchemy import create_engine
import requests
from sqlalchemy.types import NVARCHAR, Float
import sys
import cv2
import smtplib
from email.mime.text import MIMEText
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
def v_code():
    img = driver.find_element_by_id('vfcode')
    left = int(img.location['x'])  # 获取图片左上角坐标x
    top = int(img.location['y'])  # 获取图片左上角y
    right = int(img.location['x'] + img.size['width'])  # 获取图片右下角x
    bottom = int(img.location['y'] + img.size['height'])
    # print(left,right,top,bottom)
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
    try:
        driver.get(url)
        driver.find_element_by_id('ext-comp-1008').click()
        driver.find_element_by_xpath("//div[@class='x-combo-list-inner']/div[@class='x-combo-list-item'][1]").click()
    except:
        return 0
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
def time_scope(start_t, days,post_url,cir):
    global header,num_list
    start_date = datetime.datetime.strptime(start_t, '%Y-%m-%d')
    for j in range(0, days):
        to_date = start_date + datetime.timedelta(days=j)
        # 每天置零
        x = 0
        col = ['主叫号码', '被叫号码', '起始时间', '通话时长', '通话费用', '套餐时长', '套餐费用', '通话类型', '计费方式']
        df = pd.DataFrame(columns=col)
        for i in range(9, 12, 1):
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
        #修改数据库的表,存到两张表中
        df.to_sql('phone_record', engine, if_exists='append', index=False, dtype=dtypedict)
        df.to_sql('phone_record_mid1', engine, if_exists='append', index=False, dtype=dtypedict)
        num_list.append(str(to_date.date())+':'+str(df.shape[0]))
        print(cir,i, '运营商有%d行'%df.shape[0])
if __name__ == '__main__':
    conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
    engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
    driver = webdriver.Chrome()
    #各项目数量list
    num_list = []
    list2 = ['运营商', '网址', '账号', '密码','防护网','密码2','Referer','post_url']
    df2 = pd.DataFrame(columns=list2)
    df2.loc[0] = ['星河', 'http://39.98.109.6:9090', '深圳臻信', 'yzWv3RAz','http://39.98.109.6:5101/', 'shenz523','http://39.98.109.6:9090/customer/chs/query-cus-cdr.html','http://39.98.109.6:9090/customer/WebGetCustomerCdr']
    df2.loc[1] = ['金枝', 'http://112.74.44.161:9090/customer/chs/index.html', '易联达', '888888','http://112.74.44.161:5101/','KEN421pzh$@!','http://112.74.44.161:9090/customer/chs/query-cus-cdr.html','http://112.74.44.161:9090/customer/WebGetCustomerCdr']
    df2.loc[2] = ['众信','http://47.107.103.58:9090/chs/','臻信催收61151398','888888','','','http://47.107.103.58:9090/chs/query-cus-cdr.html','http://47.107.103.58:9090/chs/query-cus-cdr.jsp']
    df2.loc[3] = ['小码', 'http://120.79.30.202:7348/customer/chs/index.html', '深圳市云诺信科技全通催收', 'mjjgKGsG', '', '', 'http://120.79.30.202:7348/customer/chs/query-cus-cdr.html','http://120.79.30.202:7348/customer/WebGetCustomerCdr']
    df2.loc[4] = ['京蓝宇', 'http://60.10.163.120:3459/chs/', '深圳融臻催收', 'JIsmVpkT','','','http://60.10.163.120:3459/chs/query-cus-cdr.html','http://60.10.163.120:3459/chs/query-cus-cdr.jsp']
    for i in df2.itertuples():
        if i.运营商 == '星河' or i.运营商 == '金枝':
            try:
                driver.get(i.防护网)
                driver.find_element_by_id('passwd').send_keys(i.密码2)
                current_window = driver.current_window_handle
                driver.find_element_by_id('button').click()
                time.sleep(5)
            except:
                continue
        cookie=login(i.网址, i.账号, i.密码)
        if (cookie != 0):
            header = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
                "Cookie": cookie,
                "Referer": i.Referer
            }
            cur = conn.cursor()
            cur.execute("select left(MAX(起始时间),10) from db_danqi.dbo.phone_record_mid1 where 运营商 = %s",i.运营商)
            max_date = cur.fetchall()
            cur.close()
            #取爬取的最近日期
            date_start = str((datetime.datetime.strptime(max_date[0][0], '%Y-%m-%d') + datetime.timedelta(days=1)).date())
            n = int(str(datetime.datetime.now().date() - datetime.datetime.strptime(max_date[0][0], '%Y-%m-%d').date())[0:2].replace(':',''))
            # # # date_start=str(datetime.datetime.now().date()-datetime.timedelta(days=1))
            # print( max_date ,datetime.datetime.now().date(),n)
            print("开始爬取%s,起始时间:%s,爬取%d天"%(i.运营商,date_start,n))
            time_scope(date_start,n,i.post_url,i.运营商)
            # time_scope('2021-09-09',17,i.post_url,i.运营商)
    mail_text = '\n'
    # for i in range(0, len(num_list)):
    #     mail_text = mail_text + df2.iloc[i%5, 0] + '爬取的数量：' + num_list[i] + '\n'
    # tit = date_start + '爬取成功'+'上午'
    # send_mail(tit, mail_text)
    conn.close()
    driver.close()
    driver.quit()



