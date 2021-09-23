import pymssql
from sqlalchemy import create_engine
import win32com.client
from openpyxl  import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from datetime import datetime
conn = pymssql.connect(server="192.168.10.9", user="sj_xudq", password="oQJhkk#53", database="db_danqi")
engine = create_engine('mssql+pymssql://sj_xudq:oQJhkk#53@192.168.10.9/db_danqi')
cur = conn.cursor()
cur.execute("exec db_danqi.dbo.M3_360_p")
conn.commit()
cur.close()
conn.close()
detail=pd.read_sql("select * from db_danqi.dbo.M3_360_1",con=engine)
print(detail)
num = detail.shape[0]
print("已执行完存储过程")
path = r"360-M3-模板.xlsx"
wb=load_workbook(path)
wb.remove(wb['Sheet1'])
wb.create_sheet('Sheet1',index=None)

for row in dataframe_to_rows(detail,index=False):
	wb['Sheet1'].append(row)
wb.save(path)
wb.close()

# sleep(5)
xlApp = win32com.client.Dispatch("Excel.Application")
xlApp.Visible = False
import os
path_ab =  str(os.path.abspath(os.curdir))+'\\'
xlBook = xlApp.Workbooks.Open(path_ab+'360-M3-模板.xlsx')
# xlBook = xlApp.Workbooks.Open(r"360-M3-模板.xlsx")
xlBook.Save()
xlBook.Close()

df=pd.read_excel(path,sheet_name='案件导入模板')
df=df[0:num]
df['卡号']=df['卡号'].astype('str')
df['证件号']=df['证件号'].astype('str')

now = datetime.now().date()
this_month_start = str(datetime(now.year, now.month, 1).date()).replace("-","")
#当月的批次文件名
path2 = "QHJR-"+this_month_start+"-M3.xls"
df1=pd.read_excel(path2)
df1['个案序列号'][df1['个案序列号'].notnull()]=df1['个案序列号'][df1['个案序列号'].notnull()].astype(str).replace('.*','',regex=True)

df2=pd.concat([df1,df],axis=0)
df2['卡号']=df2['卡号'].astype('str')
df2['证件号']=df2['证件号'].astype('str')
df2.to_excel(path2 ,index=None)
print("运行完成")
a=input()
# from openpyxl import load_workbook
# wb = load_workbook(r"C:\Users\Administrator\Desktop\360.xlsx",data_only=True)
# ws = wb['案件导入模板']
# rows=ws.max_row  #获取行数
# cols=ws.max_column  #获取列数
# print(rows,cols)
# print(ws['A'])
# for i in ws.values:
#     print(i)

#遍历行
# for i in ws.values:
#     print(i)

# wb.get_sheet_by_name('Sheet1')
# 实例化
# wb = Workbook()
# 激活 worksheet
# ws = wb.active


# df=pd.read_excel(r"C:\Users\Administrator\Desktop\360-M3-模板-月末.xls",sheet_name=None)
# sheet_name=list(df.keys())
#
# df_1=df['Sheet1']
# df_11=df_1.drop(df_1.index[0:df_1.shape[0]])
# detail=pd.read_sql("select * from db_danqi.dbo.M3_360",con=engine)
# df['Sheet1']=df_11.append(detail)
# with pd.ExcelWriter(r"C:\Users\Administrator\Desktop\360-M3-模板-月末.xls") as writer:
#     df['Sheet1'].to_excel(writer,index=None,sheet_name='Sheet1')


