# -*- coding: utf-8 -*-
"""
Created on Dec 2018

@author: Reza
"""

import time
import pandas as pd
import datetime
from datetime import timedelta
import glob
import shutil
import pyodbc
import numpy as np
import xlsxwriter
from openpyxl import load_workbook

# Define Dates
year= datetime.date.today().year
month= datetime.date.today().month
day= datetime.date.today().day
month = datetime.date(year, month, day).strftime('%b')
TodaysDate = time.strftime("%d%m%Y")

server = 'My Server address'
database = 'Database name'
username = 'Username'
password = 'Password'
conn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
writer = pd.ExcelWriter('Client1.xlsx',engine='xlsxwriter')

'*****************GMV******************'
query = "SELECT [date], [cc],[product] ,[merchantName],[device], sum([orderValue]) as [GMV]" \
		"FROM [dbo].[ViewAllRevenue] " \
		"WHERE [merchantName] in ('Client1') and [date] BETWEEN '2019-04-01' AND '2019-07-31'" \
		"GROUP BY [date],[cc],[product],[merchantName],[device]" \
		"ORDER BY [date],[cc],[product],[merchantName],[device]"
first = pd.read_sql(query,conn)
df = pd.DataFrame(first)
df['merchantName'] = df['merchantName'].str.lower()

df['date'] = pd.to_datetime(df['date'])
df['date'] = df['date'].dt.month.map(str) + "-" + df['date'].dt.year.map(str)

df['product'] = df['product'].str.lower()
df['product'] = df['product'].replace({'cood':'coupon', 'n/a':'coupon','pc':'shop', 'static':'shop','trends':'shop','community' : 'shop', 'insights' : 'shop'})

product = np.round(pd.pivot_table(df, index=['cc','product'],columns = ['date'], values=['GMV'],aggfunc='sum'),0)
product.to_excel(writer,sheet_name='GMV',startrow=0 , startcol=0) 

all = np.round(pd.pivot_table(df,index = ['cc'],columns = ['date'], values=['GMV'],aggfunc='sum'),0)
all.to_excel(writer,sheet_name='GMV',startrow=0, startcol=9)

device = np.round(pd.pivot_table(df, index=['cc','device'],columns = ['date'], values=['GMV'],aggfunc='sum'),0)
device.to_excel(writer,sheet_name='GMV',startrow=18 , startcol=0) 
print('GMV Downloaded.')

'*****************CLICKS******************'

query = "SELECT [date], [cc], [product], [merchantName],[device],[level0Category], sum([ga_totalEvents]) as [Clicks]" \
		"FROM [dbo].[FactAnalyticsConversion]" \
		"WHERE ([merchantName] in ('Client1', 'Client2', 'Client3', 'Client4')) and [date] BETWEEN '2019-04-01' AND '2019-07-31'" \
		"GROUP BY [date], [cc], [product], [merchantName],[device],[level0Category] " \
		"ORDER BY [date], [cc], [product], [merchantName],[device],[level0Category] "
one = pd.read_sql(query,conn)
df = pd.DataFrame(one)
df['merchantName'] = df['merchantName'].str.lower()

df['date'] = pd.to_datetime(df['date'])
df['date'] = df['date'].dt.month.map(str) + "-" + df['date'].dt.year.map(str)
df = df.drop(df[(df['merchantName'] != 'client1')].index)

all = pd.pivot_table(df,index = ['cc'],columns = ['date'], values=['Clicks'],aggfunc='sum')
all.to_excel(writer,sheet_name='Clicks',startrow=0, startcol=9)

device = pd.pivot_table(df, index=['cc','device'],columns = ['date'], values=['Clicks'],aggfunc='sum')
device.to_excel(writer,sheet_name='Clicks',startrow=18 , startcol=0)

df['product'] = df['product'].str.lower()
df['product'] = df['product'].replace({'cood':'coupon', 'n/a':'coupon','pc':'shop', 'static':'shop','trends':'shop','community' : 'shop', 'insights' : 'shop'})
product = pd.pivot_table(df, index=['cc','product'],columns = ['date'], values=['Clicks'],aggfunc='sum')
product.to_excel(writer,sheet_name='Clicks',startrow=0 , startcol=0)

cat = pd.pivot_table(df, index = ['cc','date'],columns = ['level0Category'], values=['Clicks'],aggfunc='sum')
cat.to_excel(writer,sheet_name='Clicks',startrow=18 , startcol=9)

print('Clicks Downloaded.')

'*****************CLICKSHARE******************'

query = "SELECT [date], [cc], [merchantName],[device],[level0Category], sum([ga_totalEvents]) as [Clicks]"\
		"FROM [dbo].[FactAnalyticsConversion] " \
		"WHERE ([merchantName] in ('Client1', 'Client2', 'Client3', 'Client4')) and [date] BETWEEN '2019-04-01' AND '2019-07-31'" \
		"GROUP BY [date], [cc], [merchantName],[device],[level0Category]" \
		"ORDER BY [date], [cc], [merchantName],[device],[level0Category]"

first = pd.read_sql(query,conn)
df = pd.DataFrame(first)
df['merchantName'] = df['merchantName'].str.lower()
df['date'] = pd.to_datetime(df['date'])
df['date'] = df['date'].dt.month.map(str) + "-" + df['date'].dt.year.map(str)

main = df[df['merchantName']=='client1'].groupby(['cc','date','merchantName','device'])['Clicks'].sum()
comp = df[(df['merchantName']!='client1')].groupby(['cc','date','merchantName','device'])['Clicks'].sum()

final = pd.concat([main,comp],axis=1)
final.columns = ['Main', 'Comp']

all = pd.pivot_table(final,index = ['cc'],columns = ['date'], values=['Main', 'Comp'],aggfunc='sum')
all.to_excel(writer,sheet_name='Clickshare',startrow=0 , startcol=0)

device = pd.pivot_table(final, index = ['cc','date'],columns = ['device'], values=['Main', 'Comp'],aggfunc='sum')
device.to_excel(writer,sheet_name='Clickshare',startrow=11 , startcol=0)

#df = df.drop(df[(df['level0Category'] != 'fashion') & (df['level0Category'] != 'health-beauty')].index)
main = df[df['merchantName']=='client1'].groupby(['cc','date','device','level0Category'])['Clicks'].sum()
comp = df[(df['merchantName']!='client1')].groupby(['cc','date','device','level0Category'])['Clicks'].sum()

final = pd.concat([main,comp],axis=1)
final.columns = ['Main', 'Comp']

cat = pd.pivot_table(final, index = ['cc','date'],columns = ['level0Category'], values=['Main', 'Comp'],aggfunc='sum')
cat.to_excel(writer,sheet_name='Clickshare',startrow=40 , startcol=0)

print('Clickshare Downloaded.')

'*******************REVENUE/RPC****************'

query = "SELECT [date], [cc], [product], [merchantName],[device],[level0Category], sum([commission]) as [Revenue]"\
		"FROM [dbo].[ViewAllRevenue] " \
		"WHERE ([merchantName] in ('Client1', 'Client2', 'Client3', 'Client4')) and [date] BETWEEN '2019-04-01' AND '2019-07-31'" \
		"GROUP BY [date], [cc], [product], [merchantName],[device],[level0Category]" \
		"ORDER BY [date], [cc], [product], [merchantName],[device],[level0Category]"
first = pd.read_sql(query,conn)

df = one
df2 = first
df2['date'] = pd.to_datetime(df2['date'])
df2['date'] = df2['date'].dt.month.map(str) + "-" + df2['date'].dt.year.map(str)
df2['merchantName'] = df2['merchantName'].str.lower() 

s = pd.pivot_table(df, index=['date','cc','merchantName'],values=['Clicks'],aggfunc='sum')
u = pd.pivot_table(df2, index=['date','cc','merchantName'],values=['Revenue'],aggfunc='sum')
u['Revenue'] = u['Revenue'] - (u['Revenue']*13)/100
m = u['Revenue']/s['Clicks'] 

result = pd.concat([s,u],axis=1)
result = pd.concat([result,m],axis=1)
result.columns = ['Clicks', 'Revenue','RPC']

result = np.round(pd.pivot_table(result,index = ['cc','date'], columns = ['merchantName'], values=['RPC'],aggfunc='sum'),3)
result.to_excel(writer,sheet_name='RPC',startrow=0 , startcol=0)

o = pd.pivot_table(df, index=[ 'date','cc','merchantName','level0Category','device'],values=['Clicks'],aggfunc='sum')
n = pd.pivot_table(df2, index=['date','cc','merchantName','level0Category','device'],values=['Revenue'],aggfunc='sum')
n['Revenue'] = n['Revenue'] - (n['Revenue']*13)/100
e = n['Revenue']/o['Clicks']

result = pd.concat([o,n],axis=1)
last = pd.concat([result,e],axis=1)
last.columns = ['Clicks','Revenue','RPC']

cat = np.round(pd.pivot_table(last,index = ['cc','date'], columns = ['level0Category'], values=['RPC'],aggfunc='sum'),3)
cat.to_excel(writer,sheet_name='RPC',startrow=0 , startcol=7)

devic = np.round(pd.pivot_table(last,index = ['cc','device'], columns = ['date'], values=['RPC'],aggfunc='sum'),3)
devic.to_excel(writer,sheet_name='RPC',startrow=32 , startcol=0)

devic2 = np.round(pd.pivot_table(last,index = ['cc','merchantName','device'], columns = ['date'], values=['RPC'],aggfunc='sum'),3)
devic2.to_excel(writer,sheet_name='RPC',startrow=32 , startcol=7)

print('RPC Downloaded.')

'****************ORDER/CONVERSIONRATE*******************'
query = "SELECT [date], [cc], [product], [merchantName],[device],[level0Category], sum([order]) as [Order]"\
		"FROM [dbo].[ViewAllRevenue] " \
		"WHERE ([merchantName] in ('Client1', 'Client2', 'Client3', 'Client4')) and [date] BETWEEN '2019-04-01' AND '2019-07-31'" \
		"GROUP BY [date], [cc], [product], [merchantName],[device],[level0Category]" \
		"ORDER BY [date], [cc], [product], [merchantName],[device],[level0Category]"
extract = pd.read_sql(query,conn)
first = extract
first['merchantName'] = first['merchantName'].str.lower()

first = first.drop(first[(first['merchantName'] != 'client1')].index)
first['date'] = pd.to_datetime(first['date'])
first['date'] = first['date'].dt.month.map(str) + "-" + first['date'].dt.year.map(str)

all = pd.pivot_table(first,index = ['cc'], columns = ['date'], values=['Order'],aggfunc='sum')
all.to_excel(writer,sheet_name='Order',startrow=0, startcol=10)

device = pd.pivot_table(first, index=['cc','device'],columns = ['date'], values=['Order'],aggfunc='sum')
device.to_excel(writer,sheet_name='Order',startrow=18, startcol=0)

first['product'] = first['product'].str.lower()
first['product'] = first['product'].replace({'cood':'coupon', 'n/a':'coupon','nan':'coupon','pc':'shop', 'static':'shop','trends':'shop','ff':'shop'})
product = pd.pivot_table(first, index=['cc','product'],columns = ['date'], values=['Order'],aggfunc='sum')
product.to_excel(writer,sheet_name='Order',startrow=0 , startcol=0)

cat = pd.pivot_table(first, index = ['cc','date'],columns = ['level0Category'], values=['Order'],aggfunc='sum')
cat.to_excel(writer,sheet_name='Order',startrow=18 , startcol=10)

print('Order Downloaded.')

df = one
df['product'] = df['product'].str.lower()
df['product'] = df['product'].replace({'cood':'coupon', 'n/a':'coupon','pc':'shop', 'static':'shop','trends':'shop','community':'shop', 'insights' : 'shop'})
df2 = extract
df2['date'] = pd.to_datetime(df2['date'])
df2['date'] = df2['date'].dt.month.map(str) + "-" + df2['date'].dt.year.map(str)
df2['merchantName'] = df2['merchantName'].str.lower()
df2['product'] = df2['product'].str.lower()
df2['product'] = df2['product'].replace({'cood':'coupon', 'n/a':'coupon','nan':'coupon','pc':'shop', 'static':'shop','trends':'shop','community':'shop', 'insights' : 'shop','ff':'shop'})

o = pd.pivot_table(df, index=[ 'date','cc','merchantName'],values=['Clicks'],aggfunc='sum')
n = pd.pivot_table(df2, index=['date','cc','merchantName'],values=['Order'],aggfunc='sum')
e = (n['Order']/o['Clicks'])*100

result = pd.concat([o,n],axis=1)
last = pd.concat([result,e],axis=1)
last.columns = ['Clicks','Order','Conversion Rate']

merchant = np.round(pd.pivot_table(last, index=['cc','merchantName'],columns = ['date'], values=['Conversion Rate'],aggfunc='sum'),2)
merchant.to_excel(writer,sheet_name='Conversion Rate',startrow=0 , startcol=0)

o = pd.pivot_table(df, index=[ 'date','cc','merchantName','product'],values=['Clicks'],aggfunc='sum')
n = pd.pivot_table(df2, index=['date','cc','merchantName','product'],values=['Order'],aggfunc='sum')
e = (n['Order']/o['Clicks'])*100

result = pd.concat([o,n],axis=1)
last = pd.concat([result,e],axis=1)
last.columns = ['Clicks','Order','Conversion Rate']

product = np.round(pd.pivot_table(last, index=['cc','merchantName'],columns = ['date','product'], values=['Conversion Rate'],aggfunc='sum'),2)
product.to_excel(writer,sheet_name='Conversion Rate',startrow=0 , startcol=8)

df = df.drop(df[(df['merchantName'] != 'client1')].index)
df2 = df2.drop(df2[(df2['merchantName'] != 'client1')].index)

o = pd.pivot_table(df, index=[ 'date','cc','merchantName','device'],values=['Clicks'],aggfunc='sum')

n = pd.pivot_table(df2, index=['date','cc','merchantName','device'],values=['Order'],aggfunc='sum')

e = (n['Order']/o['Clicks'])*100

result = pd.concat([o,n],axis=1)
last = pd.concat([result,e],axis=1)
last.columns = ['Clicks','Order','Conversion Rate']

device = np.round(pd.pivot_table(last, index=['cc','device'],columns = ['date'], values=['Conversion Rate'],aggfunc='sum'),2)
device.to_excel(writer,sheet_name='Conversion Rate',startrow=30, startcol=0)

o = pd.pivot_table(df, index=[ 'date','cc','merchantName','level0Category'],values=['Clicks'],aggfunc='sum')
n = pd.pivot_table(df2, index=['date','cc','merchantName','level0Category'],values=['Order'],aggfunc='sum')
n['Order'] = n['Order']*100
e = n['Order']/o['Clicks']

result = pd.concat([o,n],axis=1)
last = pd.concat([result,e],axis=1)
last.columns = ['Clicks','Order','Conversion Rate']

cat = np.round(pd.pivot_table(last,index=['cc','date'],columns = ['level0Category'], values=['Conversion Rate'],aggfunc='sum'),2)
cat.to_excel(writer,sheet_name='Conversion Rate',startrow=50 , startcol=0) 

print('Conversion Rate Downloaded.')

writer.save()

