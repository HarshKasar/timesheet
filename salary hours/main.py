import imp
from pickle import TRUE
import csv
from flask import Flask,jsonify
from openpyxl import load_workbook
from pip import main
app = Flask(__name__)
@app.route('/')
def Helloworld():
    return 'Hello World!'
@app.route('/con')
def convert():
    book= load_workbook("timesheet2.xlsx")
    sheet=book.active
    d=30
    a=[]
    n=[]
    e=[]
    tim=[]
    dic1={}
    date=[]
    for h in sheet['a4':'a55']:
        for j in h:
            e.append(j.value)
    for row in sheet['s4':'s55']:
        for cell in row:
            b=(cell.value)
            a.append(b)
    #print(a)
    for r in sheet['n4':'n55']:
        for c in r:
            n.append(c.value)
    for t in a:
        dic={'hrs':t}
        tim.append(dic)
    for d in n:
       dic1={'date':d}
       date.append(dic1)
    #print(n)
    h=list(zip(a,n))

    fields_name=['Empid','hrs','date']
    time=[
        #{'Empid':e[0],'hrs':a[0],'date':n[0]},
        #{'Empid':e[1],'hrs':a[1],'date':n[1]},
        #{'Empid':e[2],'hrs':a[2],'date':n[2]},
        #{'Empid':e[3],'hrs':a[3],'date':n[3]},
        #{'Empid':e[4],'hrs':a[4],'date':n[4]},
        #{'Empid':e[5],'hrs':a[5],'date':n[5]},
        #{'Empid':e[6],'hrs':a[6],'date':n[6]},
        #{'Empid':e[7],'hrs':a[7],'date':n[7]},
        #{'Empid':e[8],'hrs':a[8],'date':n[8]},
        #{'Empid':e[9],'hrs':a[9],'date':n[9]},
        #{'Empid':e[10],'hrs':a[10],'date':n[10]},
        #{'Empid':e[11],'hrs':a[11],'date':n[11]},
        #{'Empid':e[12],'hrs':a[12],'date':n[12]},
        #{'Empid':e[13],'hrs':a[13],'date':n[13]},
        #{'Empid':e[14],'hrs':a[14],'date':n[14]},
        #{'Empid':e[15],'hrs':a[15],'date':n[15]},
        #{'Empid':e[16],'hrs':a[16],'date':n[16]},
        #{'Empid':e[17],'hrs':a[17],'date':n[17]},
        #{'Empid':e[18],'hrs':a[18],'date':n[18]},
        #{'Empid':e[19],'hrs':a[19],'date':n[19]},
        #{'Empid':e[20],'hrs':a[20],'date':n[20]},
        #{'Empid':e[21],'hrs':a[21],'date':n[21]},
        #{'Empid':e[22],'hrs':a[22],'date':n[22]},
        #{'Empid':e[23],'hrs':a[23],'date':n[23]},
        #{'Empid':e[24],'hrs':a[24],'date':n[24]},
        #{'Empid':e[25],'hrs':a[25],'date':n[25]},
        #{'Empid':e[26],'hrs':a[26],'date':n[26]},
        #{'Empid':e[27],'hrs':a[27],'date':n[27]},
        #{'Empid':e[28],'hrs':a[28],'date':n[28]},
        #{'Empid':e[29],'hrs':a[29],'date':n[29]},
        #{'Empid':e[30],'hrs':a[30],'date':n[30]},
        #{'Empid':e[31],'hrs':a[31],'date':n[31]},
        #{'Empid':e[32],'hrs':a[32],'date':n[32]},
        #{'Empid':e[33],'hrs':a[33],'date':n[33]},
        #{'Empid':e[34],'hrs':a[34],'date':n[34]},
        #{'Empid':e[35],'hrs':a[35],'date':n[35]},
        #{'Empid':e[36],'hrs':a[36],'date':n[36]},
        #{'Empid':e[37],'hrs':a[37],'date':n[37]},
        #{'Empid':e[38],'hrs':a[38],'date':n[38]},
        #{'Empid':e[39],'hrs':a[39],'date':n[39]},
        #{'Empid':e[40],'hrs':a[40],'date':n[40]},
        #{'Empid':e[41],'hrs':a[41],'date':n[41]},
        #{'Empid':e[42],'hrs':a[42],'date':n[42]},
        #{'Empid':e[43],'hrs':a[43],'date':n[43]},
        #{'Empid':e[44],'hrs':a[44],'date':n[44]},
        #{'Empid':e[45],'hrs':a[45],'date':n[45]},
        #{'Empid':e[46],'hrs':a[46],'date':n[46]},
        #{'Empid':e[47],'hrs':a[47],'date':n[47]},
        #{'Empid':e[48],'hrs':a[48],'date':n[48]},
        #{'Empid':e[49],'hrs':a[49],'date':n[49]},
        #{'Empid':e[50],'hrs':a[50],'date':n[50]},
    ]
    with open('time.csv','w') as csvfile:
        writer=csv.writer(csvfile)
        writer.writerows(h)  
    return jsonify(h)
if __name__=="__main__":
    app.run(debug=TRUE)
