import time, requests, math, json
# driver.find_element_by_class_name('x-btn-text').click()
import datetime
import pandas as pd
import pymssql
from sqlalchemy import create_engine
from sqlalchemy.types import NVARCHAR, Float, Integer
import cv2
import smtplib
from email.mime.text import MIMEText
# from selenium import webdriver
# driver=webdriver.Chrome()pu
# driver.get("https://cdr.csm.zthltx.com:8709/login?redirect=%2Fbalance%2Findex")
# driver.find_element_by_name('companyName').clear()
# driver.find_element_by_name('companyName').send_keys('深圳融臻资产惠州分公司')
# driver.find_element_by_name('accountId').clear()
# driver.find_element_by_name('accountId').send_keys('shenzhenrongzhenzichanhuizhoufengongsi')
# driver.find_element_by_name('password').clear()
# driver.find_element_by_name('password').send_keys('Cfs2020')
# driver.find_element_by_class_name('el-button.el-button--primary.el-button--default').click()
def send_mail(titles,contexts):
    mail_host = 'smtp.163.com'
    mail_user = 'dq1601430804'
    mail_pass = 'NFWHVBPNKTPLNPEC'
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
def date_scope(start_t, days):
    global response, header
    # 天数
    for day in range(0, days):
        start_date = datetime.datetime.strptime(start_t, '%Y-%m-%d').date()
        to_start_date = start_date + datetime.timedelta(days=day)
        #
        url = "https://cdr.csm.zthltx.com:8709/api/customer/call/record?page={n}&pageSize=1000&startTime={s_time}+00:00:00&endTime={e_time}+23:59:59&sortField=&order=".format(n=1, s_time=to_start_date, e_time=to_start_date)
        responses = requests.get(url=url, headers=header)
        page_number = math.ceil(int(responses.json().get('data').get('total')) / 1000) + 1
        col = ['线路号码', '客户号码', '开始时间', '挂断时间', '通话时长', '计费次数', '费用']
        df = pd.DataFrame(columns=col)
        # print('df的值',df.shape[0])
        x=0
        # 页数
        for n in range(1, page_number):
            #
            url = "https://cdr.csm.zthltx.com:8709/api/customer/call/record?page={n}&pageSize=1000&startTime={s_time}+00:00:00&endTime={e_time}+23:59:59&sortField=&order=".format(n=n, s_time=to_start_date, e_time=to_start_date)
            responses = requests.get(url=url, headers=header)
            call_records = responses.json().get('data').get('list')
            for item in call_records:
                orilinenumber = item.get('orilinenumber')
                oriphone = item.get('oriphone')
                begintime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(item.get('begintime') / 1000))
                hanguptime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(item.get('hanguptime') / 1000))
                callduration = (item.get('callduration')) * 1.0 / 1000
                billingNumber = item.get('billingNumber')
                totalfee = item.get('totalfee')
                df.loc[x] = [orilinenumber, oriphone, begintime, hanguptime, callduration, billingNumber, totalfee]
                x += 1
            print('第%d页'%n)
        df['运营商'] = '长喜'
        df['插入时间'] = str(datetime.datetime.now())
        dtypedict = {
            '线路号码': NVARCHAR(20), '客户号码':NVARCHAR(20), '开始时间': NVARCHAR(30),'插入时间':NVARCHAR(30),
            '挂断时间': NVARCHAR(30), '通话时长': Float, '计费次数': Integer, '费用': Float, '运营商': NVARCHAR(10)
        }
        # df.to_sql('phone_record2_mid1', engine, if_exists='append', index=False, dtype=dtypedict)
        df.to_sql('phone_record2', engine, if_exists='append', index=False, dtype=dtypedict)
        print('第%d天'%day)
        mail_text = '\n' + '长喜爬取的数量：' + str(df.shape[0]) + '\n'
        tit = str(start_date) + '上午爬取成功'
        send_mail(tit, mail_text)
if __name__ == '__main__':
    URL = "https://cdr.csm.zthltx.com:8709/api/login"
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36"
    }
    from_data = {
        "accountId": "shenzhenrongzhenzichanhuizhoufengongsi",
        "password": "46a3f7d9245a2e268740a0471e6be26b",
        "companyName": "深圳融臻资产惠州分公司"
    }
    response = requests.post(url=URL, headers=header, data=from_data)
    header['Cookie'] = response.cookies.items()[0][0] + "=" + response.cookies.items()[0][1]
    header['token'] = response.json().get('data').get('token')

    conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
    engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
    cur = conn.cursor()
    yesterday = str((datetime.datetime.now() - datetime.timedelta(days=1)).date())
    print(yesterday)
    cur.execute("delete from db_danqi.dbo.phone_record2 where convert(varchar(10),开始时间,120)=%s",yesterday)
    conn.commit()
    print("删除完毕")
    cur.execute("select left(MAX(开始时间),10) 最大时间 from db_danqi.dbo.phone_record2")
    max_date = cur.fetchall()
    # 取爬取的最近日期

    date_start = str((datetime.datetime.strptime(max_date[0][0], '%Y-%m-%d') + datetime.timedelta(days=1)).date())
    n = int(str(datetime.datetime.now().date() - datetime.datetime.strptime(max_date[0][0], '%Y-%m-%d').date())[0:2].replace(':', ''))
    print("起始时间:%s,爬取%d天"%(date_start, n))
    # 修改时间
    # date_scope('2021-09-08', 1)
    date_scope(date_start, n)
    cur.close()
    conn.close()
    # driver.close()
    # driver.quit()


