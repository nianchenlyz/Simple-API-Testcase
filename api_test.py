#encoding:utf-8
import ConfigParser
import os
import xlrd
import re
import httplib
import urllib
from urlparse import urlparse
import json
import time
import unittest
import requests
from xlrd import open_workbook
from xlutils.copy import copy
#import pdf

currentdir=os.path.abspath(os.curdir)
def getexcel():
    casefile=currentdir + '/case.xlsx'
    if ((os.path.exists(casefile))==False):
        print "当前路径下没有case.xls，请检查！"
    data=xlrd.open_workbook(casefile)
    table = data.sheet_by_name('login')
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    for rownum in range(1,nrows):
        for col in range (3, ncols):
            value=table.cell(rownum,col).value
            if (col==3):
                method=value
            if (col==4):
                url=value
    #print table,nrows,ncols
    return table,nrows,ncols

def getexceldetail(table,row,ncols):
    li=[]
    data={'ID':0,'method':'get','url':'http://www.baidu.com'}
    for onerow in range(1, row):
        for col in range (0, ncols):
            value=table.cell(onerow,col).value
            #print value
            if (col==0):
                data['ID']=value
                #print data
            if (col==3):
                data['method']=value
            if (col==4):
                data['url']=value
            if col==ncols-1:
                data_string = json.dumps(data)
                li.append(data_string)
    return li

def request_get(url,ID):
    result={'ID':0,'code':0,'result':'','porf':''}
    result['ID']=ID
    header={'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.93 Safari/537.36'}
    res=requests.get(url,headers=header,timeout=10)
    #print type(res)
    result['code']=res.status_code
    #print res.headers
    result['result']= res.text
    if result['code']==200:
        result['porf']='Pass'
    else:
        result['porf']='Failure'
    #print result
    return result

def request_post(url,ID):
    result={'ID':0,'code':0,'result':'','porf':''}
    result['ID']=ID
    header={'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.93 Safari/537.36'}
    res=requests.post(url)
    result['code']=res.status_code
    #print res.headers
    result['result']= res.text
    if result['code']==200:
        result['porf']='Pass'
    else:
        result['porf']='Failure'
    #print result
    return result

def write_xls(table,response):
    rb = open_workbook(table)
    rs=rb.sheet_by_index(0)
    wb=copy(rb)
    ws=wb.get_sheet(0)
    if response['ID']!=0:
        row=int(response['ID'])
        col=5
        value=response['code']
        ws.write(row,col,value)
        ws.write(row,col+1,response['porf'])
        wb.save(table)
        #print table.cell(row,col) 

def main():
    (table,nrows,ncols)=getexcel()
    #print table
    #print nrows
    #print ncols
    sheet_list=getexceldetail(table,nrows,ncols)
    table='case.xlsx'
    for sheet in sheet_list:
        rr=eval(sheet)
        sheet=rr
        if sheet['method']=='GET':
            response=request_get(sheet['url'],sheet['ID'])
            write_xls(table,response)
        else:
            response=request_post(sheet['url'],sheet['ID'])
            write_xls(table,response)

if __name__ == "__main__": 
    main()
