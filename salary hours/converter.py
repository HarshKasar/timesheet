from calendar import c
import csv
#from datetime import datetime
from traceback import print_last
from openpyxl import load_workbook
import pandas as pandasForSortingCSV
import pandas as pd

book= load_workbook("Timesheet2.xlsx")
sheet=book.active
d=30
a=[]
n=[]
k=5
e=[]
dic={}
dic1={}
tim=[]
date=[]
#taking empid
for h in sheet['a4':'a34']:
    for j in h:
        e.append({'Empid':j.value})
#taking working hrs
for row in sheet['s4':'s34']:
    for cell in row:
        b=(cell.value)
        a.append(b)
#print(a)
#taking dates
for r in sheet['n4':'n34']:
    for c in r:
        n.append(c.value)
for t in a:
    dic={'hrs':t}
    tim.append(dic)
for d in n:
    dic1={'date':d}
    date.append(dic1)
print(tim)
print(date)
#print(n)
#print(dic)
#for k in dic.keys():
#if k in dic.keys():
 #   print()
fields_name=['Empid','hrs','date']
time=[
#    {'No.':1,'Empid':e[0],'hrs':a[0],'date':n[0]},
#    {'No.':2,'Empid':e[1],'hrs':a[1],'date':n[1]},
#    {'No.':3,'Empid':e[2],'hrs':a[2],'date':n[2]},
#    {'No.':4,'Empid':e[3],'hrs':a[3],'date':n[3]},
#    {'No.':4,'Empid':e[3],'hrs':a[4],'date':n[4]},
#    {'No.':6,'Empid':e[0],'hrs':a[5],'date':n[5]},
#    {'No.':7,'Empid':e[0],'hrs':a[6],'date':n[6]},
#    {'No.':8,'Empid':e[0],'hrs':a[7],'date':n[7]},
#    {'No.':9,'Empid':e[0],'hrs':a[8],'date':n[8]},
#    {'No.':10,'Empid':e[0],'hrs':a[9],'date':n[9]},
#    {'No.':11,'Empid':e[0],'hrs':a[10],'date':n[10]},
#    {'No.':12,'Empid':e[0],'hrs':a[11],'date':n[11]},
#    {'No.':13,'Empid':e[0],'hrs':a[12],'date':n[12]},
#    {'No.':14,'Empid':e[0],'hrs':a[13],'date':n[13]},
#    {'No.':15,'Empid':e[0],'hrs':a[14],'date':n[14]},
#    {'No.':16,'Empid':e[0],'hrs':a[15],'date':n[15]},
#    {'No.':17,'Empid':e[0],'hrs':a[16],'date':n[16]},
#    {'No.':18,'Empid':e[0],'hrs':a[17],'date':n[17]},
#    {'No.':19,'Empid':e[0],'hrs':a[18],'date':n[18]},
#    {'No.':20,'Empid':e[0],'hrs':a[19],'date':n[19]},
#    {'No.':21,'Empid':e[0],'hrs':a[20],'date':n[20]},
#    {'No.':22,'Empid':e[0],'hrs':a[21],'date':n[21]},
#    {'No.':23,'Empid':e[0],'hrs':a[22],'date':n[22]},
#    {'No.':24,'Empid':e[0],'hrs':a[23],'date':n[23]},
#    {'No.':24,'Empid':e[0],'hrs':a[24],'date':n[24]},
#    {'No.':25,'Empid':e[0],'hrs':a[25],'date':n[25]},
#    {'No.':26,'Empid':e[0],'hrs':a[26],'date':n[26]},
#    {'No.':27,'Empid':e[0],'hrs':a[27],'date':n[27]},
#    {'No.':27,'Empid':e[0],'hrs':a[28],'date':n[28]},
#    {'No.':27,'Empid':e[0],'hrs':a[29],'date':n[29]},
#    {'No.':27,'Empid':e[0],'hrs':a[30],'date':n[30]},
]
#for i in a:
#time=[i]
with open('time.csv','w') as csvfile:
    writer=csv.DictWriter(csvfile,fieldnames=fields_name)
    writer.writeheader()
    wr=writer.writerows(tim)  
    #writer.writerows(date)   
    #writer.writerows(e)
with open('time.csv','w') as csvfile:
    writer=csv.DictWriter(csvfile,fieldnames=fields_name)
    writer.writeheader(wr)
    writer.writerows(date)
    writer.writerow()
#csvData = pandasForSortingCSV.read_csv("time.csv")
#csvData.sort_index(["hrs"],axis=0,ascending=False,inplace=True)
#print(csvData)