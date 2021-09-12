from tkinter import *
import pandas as pd
from csv import DictReader
import numpy as np
import datetime
from datetime import date, timedelta
from datetime import datetime
import decimal
from datetime import *
import shutil
import csv
import os
import pdfkit
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt
import os
import openpyxl
import openpyxl.worksheet.worksheet
from win32com.client import Dispatch


op=True
large_font = ('Verdana',30)
x = True
List=[]
s=0
y=0.0
z=0
df=''
master = Tk()
f1 = 2.5
f2 = 2.5
f3 = 2.5
f4 = 2.5
f5 = 1.7
f6 = 2
f7 = 2.5
f8 = 1.7
f9 = 1.7
f10 = 2.5
f11 = 5.0
master.geometry('5000x5000')
def paaji():
    file_name = "C:/Users/DELL/Downloads/S/print.xlsx"
    df1000 = pd.read_excel(file_name) #Read Excel file as a DataFrame
    df1000.set_index("दिनांक", inplace = True)
    wb_obj = openpyxl.load_workbook('print.xlsx')
            
    sheet_obj = wb_obj.active
    ab = df1000['कुल'].sum()
    print(ab)
    ac = df1000['एडवांस'].sum()
    c3 = sheet_obj['Z2']
    c3.value = ab
    c4 = sheet_obj['Z3']
    c4.value = ac
    c4 = sheet_obj['AA3']
    c4.value = ab-ac
    c5 = sheet_obj['AB3']
    c5.value = sum2+TVP-ad2
    wb_obj.save('print.xlsx')
    df9000 = pd.read_excel('C:/Users/DELL/Downloads/S/print.xlsx')
    df9001 = df9000.T
    df9001.to_excel('C:/Users/DELL/Downloads/S/print.xlsx')
    excel = Dispatch('Excel.Application')
    wb = excel.Workbooks.Open("C:/Users/DELL/Downloads/S/print.xlsx")

    #Activate second sheet
    excel.Worksheets(1).Activate()

    #Autofit column in active sheet
    excel.ActiveSheet.Columns.AutoFit()
    wb.Save()

    wb.Close()
    os.startfile("C:/Users/DELL/Downloads/S/print.xlsx",'print')
def Sum():
    Label(master, text='आरंभ दिनांक').grid(row=1,column=7)
    Label(master, text='समाप्ति दिनांक').grid(row=3,column=7)
    global bd
    global ed
    global bb
    bb = 'jan-27-21'
    bd = Entry(master)
    ed = Entry(master)
    bd.grid(row=1, column=8)
    ed.grid(row=3, column=8)
    pp = 0
    EN = ''
    D = ''
    X = 0.0
    def search():
        data2 = pd.read_csv("Records.csv")
        data = pd.read_csv("Records.csv")
        data.set_index("Date", inplace = True)
        df82 = data.loc[bd.get():ed.get(),['Haudi Katai Paddy(40Kg,30Kg)','Rate','Haudi Katai Rice(60Kg,50Kg)','Rate2','Paddy Stack(40Kg,30Kg)','Rate3','Doubling','Rate4','Rice Loading(60kg,50kg,40kg)','Rate5','Rice Loading(25kg)','Rate6','Polish Loading','Rate7','Rice Dhala','Rate8','Bundle Stack/Loading','Rate9','Rice Stack(25Kg)','Rate10','Rice Stack(50Kg,60Kg)','Rate11','Adv']]
        df2 = df82['Adv']
        ok =df82['Haudi Katai Paddy(40Kg,30Kg)']
        ok2 = df82['Rate']
        pro1 = ok*ok2
        ok3 = df82['Haudi Katai Rice(60Kg,50Kg)']
        ok4 = df82['Rate2']
        pro2 = ok3*ok4
        ok5 = df82['Paddy Stack(40Kg,30Kg)']
        ok6 = df82['Rate3']
        pro3 = ok5*ok6
        ok7 = df82['Rice Loading(60kg,50kg,40kg)']
        ok8 = df82['Rate5']
        pro4 = ok7*ok8
        ok9 = df82['Rice Loading(25kg)']
        ok10 = df82['Rate6']
        pro5 = ok9*ok10
        ok11 = df82['Polish Loading']
        ok12 = df82['Rate7']
        pro6 = ok11*ok12
        ok13 = df82['Rice Dhala']
        ok14 = df82['Rate8']
        pro7 = ok13*ok14
        ok15 = df82['Bundle Stack/Loading']
        ok16 = df82['Rate9']
        pro8 = ok15*ok16
        ok17 = df82['Rice Stack(25Kg)']
        ok18 = df82['Rate10']
        pro9 = ok17*ok18
        ok19 = df82['Rice Stack(50Kg,60Kg)']
        ok20 = df82['Rate11']
        ok21 = df82['Doubling']
        ok22 = df82['Rate4']
        pro10 = ok19*ok20
        pro11 = ok21*ok22
        result = pro1+pro2+pro3+pro4+pro5+pro6+pro7+pro8+pro9+pro10+pro11
        df82.to_excel('print.xlsx')
        def Excel3():
            col_names = ['दिनांक','हौदी कटाई धान(40Kg,30Kg)','Rate1','हौदी कटाई चावल(60Kg,50Kg)','Rate2','धान स्टैक(40Kg,30Kg)','Rate3','डब्लिंग','Rate4','चावल लोड(60kg,50kg,40kg)','Rate5','चावल लोड(25kg)','Rate6','पोलिश लोड','Rate7','चावल ढला','Rate8','बंडल स्टैक/लोड','Rate9','चावल स्टैक(25Kg)','Rate10','चावल स्टैक(50Kg,60Kg)','Rate11','एडवांस']
            file_name = "C:/Users/DELL/Downloads/S/print.xlsx"
            df1000 = pd.read_excel(file_name) #Read Excel file as a DataFrame
            df1000.columns=col_names
            df1000.set_index("दिनांक", inplace = True)
            df1000['कुल'] = df1000['हौदी कटाई धान(40Kg,30Kg)']*df1000['Rate1']+df1000['हौदी कटाई चावल(60Kg,50Kg)']*df1000['Rate2']+df1000['धान स्टैक(40Kg,30Kg)']*df1000['Rate3']+df1000['डब्लिंग']*df1000['Rate4']+df1000['चावल लोड(60kg,50kg,40kg)']*df1000['Rate5']+df1000['चावल लोड(25kg)']*df1000['Rate6']+df1000['पोलिश लोड']*df1000['Rate7']+df1000['चावल ढला']*df1000['Rate8']+df1000['बंडल स्टैक/लोड']*df1000['Rate9']+df1000['चावल स्टैक(25Kg)']*df1000['Rate10']+df1000['चावल स्टैक(50Kg,60Kg)']*df1000['Rate11']
            df1000.to_excel("C:/Users/DELL/Downloads/S/print.xlsx") 
        global button5
        button5 = Button(master, text='Print', width=25, command=lambda:[Excel3(),paaji()])
        button5.grid(row=20,column=12)
        data2 = pd.read_csv("VRecords.csv")
        data2.set_index("Date", inplace = True)
        df6 = data2.loc[bd.get():ed.get(),['Haudi Katai Paddy(40Kg,30Kg)','Rate','Haudi Katai Rice(60Kg,50Kg)','Rate2','Paddy Stack(40Kg,30Kg)','Rate3','Doubling','Rate4','Rice Loading(60kg,50kg,40kg)','Rate5','Rice Loading(25kg)','Rate6','Polish Loading','Rate7','Rice Dhala','Rate8','Bundle Stack/Loading','Rate9','Rice Stack(25Kg)','Rate10','Rice Stack(50Kg,60Kg)','Rate11']]
        okkk =df6['Haudi Katai Paddy(40Kg,30Kg)']
        okkk2 = df6['Rate']
        prooo1 = okkk*okkk2
        okkk3 = df6['Haudi Katai Rice(60Kg,50Kg)']
        okkk4 = df6['Rate2']
        prooo2 = okkk3*okkk4
        okkk5 = df6['Paddy Stack(40Kg,30Kg)']
        okkk6 = df6['Rate3']
        prooo3 = okkk5*okkk6
        okkk7 = df6['Rice Loading(60kg,50kg,40kg)']
        okkk8 = df6['Rate5']
        prooo4 = okkk7*okkk8
        okkk9 = df6['Rice Loading(25kg)']
        okkk10 = df6['Rate6']
        prooo5 = okkk9*okkk10
        okkk11 = df6['Polish Loading']
        okkk12 = df6['Rate7']
        prooo6 = okkk11*okkk12
        okkk13 = df6['Rice Dhala']
        okkk14 = df6['Rate8']
        prooo7 = okkk13*okkk14
        okkk15 = df6['Bundle Stack/Loading']
        okkk16 = df6['Rate9']
        prooo8 = okkk15*okkk16
        okkk17 = df6['Rice Stack(25Kg)']
        okkk18 = df6['Rate10']
        prooo9 = okkk17*okkk18
        okkk19 = df6['Rice Stack(50Kg,60Kg)']
        okkk20 = df6['Rate11']
        okkk21 = df6['Doubling']
        okkk22 = df6['Rate4']
        prooo10 = okkk19*okkk20
        prooo11 = okkk21*okkk22
        global result5
        result5 = prooo1+prooo2+prooo3+prooo4+prooo5+prooo6+prooo7+prooo8+prooo9+prooo10+prooo11
        global L5
        global L6
        global L7
        global canvas
        global scroll_y
        canvas = Canvas(master, width=400, height=200)
        scroll_y = Scrollbar(master,orient="vertical", command=canvas.yview)

        frame = Frame(canvas)
        # group of widgets
        for i in range(20):
            with pd.option_context('display.max_rows', None, 'display.max_columns', None):  
                global L
                global L2
                global L3
                global L4
                global L8
                global L9
                L=Label(master, text="Advance")
                L.grid(row=1,column=13)
                L2=Label(master, text="Sum")
                L2.grid(row=1,column=11)
                L3=Label(frame,text=df2)
                L3.grid(row=3,column=13)
                L4=Label(frame, text=result)
                L4.grid(row=3,column=11)
                L8=Label(master, text="Vardhman Payment")
                L8.grid(row=1,column=12)
                L9=Label(frame, text=result5)
                L9.grid(row=3,column=12)
        canvas.create_window(1000,1000, anchor='nw', window=frame)
        canvas.update_idletasks()

        canvas.configure(scrollregion=canvas.bbox('all'), 
                        yscrollcommand=scroll_y.set)
                        
        canvas.grid(row=2,column=12)
        scroll_y.grid(row=2,column=14)
        sum = result.sum()
        ad = df2.sum()
        va = result5.sum()
        global L10
        L5=Label(master, text=float(sum))
        L5.grid(row=5,column=11)
        L6=Label(master,text=float(ad))
        L6.grid(row=5,column=13)
        L10=Label(master, text=float(va))
        L10.grid(row=5,column=12)
        L7=Label(master, text=float(sum+va-ad))
        L7.grid(row=7,column=12)
        data6 = pd.read_csv("VRecords.csv")
        data6.set_index("Date", inplace = True)
        df8 = data6.loc[bb:ed.get(),['Haudi Katai Paddy(40Kg,30Kg)','Rate','Haudi Katai Rice(60Kg,50Kg)','Rate2','Paddy Stack(40Kg,30Kg)','Rate3','Doubling','Rate4','Rice Loading(60kg,50kg,40kg)','Rate5','Rice Loading(25kg)','Rate6','Polish Loading','Rate7','Rice Dhala','Rate8','Bundle Stack/Loading','Rate9','Rice Stack(25Kg)','Rate10','Rice Stack(50Kg,60Kg)','Rate11']]
        okkkk =df8['Haudi Katai Paddy(40Kg,30Kg)']
        okkkk2 = df8['Rate']
        proooo1 = okkkk*okkkk2
        okkkk3 = df8['Haudi Katai Rice(60Kg,50Kg)']
        okkkk4 = df8['Rate2']
        proooo2 = okkkk3*okkkk4
        okkkk5 = df8['Paddy Stack(40Kg,30Kg)']
        okkkk6 = df8['Rate3']
        proooo3 = okkkk5*okkkk6
        okkkk7 = df8['Rice Loading(60kg,50kg,40kg)']
        okkkk8 = df8['Rate5']
        proooo4 = okkkk7*okkkk8
        okkkk9 = df8['Rice Loading(25kg)']
        okkkk10 = df8['Rate6']
        proooo5 = okkkk9*okkkk10
        okkkk11 = df8['Polish Loading']
        okkkk12 = df8['Rate7']
        proooo6 = okkkk11*okkkk12
        okkkk13 = df8['Rice Dhala']
        okkkk14 = df8['Rate8']
        proooo7 = okkkk13*okkkk14
        okkkk15 = df8['Bundle Stack/Loading']
        okkkk16 = df8['Rate9']
        proooo8 = okkkk15*okkkk16
        okkkk17 = df8['Rice Stack(25Kg)']
        okkkk18 = df8['Rate10']
        proooo9 = okkkk17*okkkk18
        okkkk19 = df8['Rice Stack(50Kg,60Kg)']
        okkkk20 = df8['Rate11']
        okkkk21 = df8['Doubling']
        okkkk22 = df8['Rate4']
        proooo10 = okkkk19*okkkk20
        proooo11 = okkkk21*okkkk22
        global result8
        result8 = proooo1+proooo2+proooo3+proooo4+proooo5+proooo6+proooo7+proooo8+proooo9+proooo10+proooo11
        global TVP
        TVP = result8.sum()
        data3 = pd.read_csv("Records.csv")
        data3.set_index("Date", inplace = True)
        df5 = data3.loc[bb:ed.get(),['Haudi Katai Paddy(40Kg,30Kg)','Rate','Haudi Katai Rice(60Kg,50Kg)','Rate2','Paddy Stack(40Kg,30Kg)','Rate3','Doubling','Rate4','Rice Loading(60kg,50kg,40kg)','Rate5','Rice Loading(25kg)','Rate6','Polish Loading','Rate7','Rice Dhala','Rate8','Bundle Stack/Loading','Rate9','Rice Stack(25Kg)','Rate10','Rice Stack(50Kg,60Kg)','Rate11','Adv']]
        df6 = df5['Adv']
        okk =df5['Haudi Katai Paddy(40Kg,30Kg)']
        okk2 = df5['Rate']
        proo1 = okk*okk2
        okk3 = df5['Haudi Katai Rice(60Kg,50Kg)']
        okk4 = df5['Rate2']
        proo2 = okk3*okk4
        okk5 = df5['Paddy Stack(40Kg,30Kg)']
        okk6 = df5['Rate3']
        proo3 = okk5*okk6
        okk7 = df5['Rice Loading(60kg,50kg,40kg)']
        okk8 = df5['Rate5']
        proo4 = okk7*okk8
        okk9 = df5['Rice Loading(25kg)']
        okk10 = df5['Rate6']
        proo5 = okk9*okk10
        okk11 = df5['Polish Loading']
        okk12 = df5['Rate7']
        proo6 = okk11*okk12
        okk13 = df5['Rice Dhala']
        okk14 = df5['Rate8']
        proo7 = okk13*okk14
        okk15 = df5['Bundle Stack/Loading']
        okk16 = df5['Rate9']
        proo8 = okk15*okk16
        okk17 = df5['Rice Stack(25Kg)']
        okk18 = df5['Rate10']
        proo9 = okk17*okk18
        okk19 = df5['Rice Stack(50Kg,60Kg)']
        okk20 = df5['Rate11']
        okk21 = df5['Doubling']
        okk22 = df5['Rate4']
        proo10 = okk19*okk20
        proo11 = okk21*okk22
        result2 = proo1+proo2+proo3+proo4+proo5+proo6+proo7+proo8+proo9+proo10+proo11
        global sum2
        sum2 = result2.sum()
        global ad2
        ad2 = df6.sum()
        global L90
        L90=Label(master, text=float("{:.2f}".format(sum2+TVP-ad2)))
        L90.grid(row=8,column=12)
    def delete_label():
        L.destroy()
        L2.destroy()
        L3.destroy()
        L4.destroy()
        L5.destroy()
        L6.destroy()
        L7.destroy()
        L8.destroy()
        L9.destroy()
        L10.destroy()
        L90.destroy()
        canvas.destroy()
        scroll_y.destroy()
        bd.delete(0,END)
        ed.delete(0,END)
        button5.destroy()
    button = Button(master, text='Submit', width=25, command=search)
    button.grid(row=5,column=8)
    button2 = Button(master, text='Clear', width=25, command=delete_label)
    button2.grid(row=6,column=8)
def Vardhman():
    def some_callbackk(event):
        event.widget.delete(0, "end")
        return None
    a1 = Label(master, text='दिनांक')
    a1.grid(row=20)
    a2 = Label(master, text='हौदी कटाई धान(40Kg,30Kg)')
    a2.grid(row=21)
    ee1 = Entry(master)
    ee2 = Entry(master)
    ee2.insert(0,0)
    ee2.bind("<Button-1>", some_callbackk)
    ee1.grid(row=20, column=1)
    ee2.grid(row=21, column=1)
    a3 = Label(master, text='हौदी कटाई चावल(60Kg,50Kg)')
    a3.grid(row=22)
    a4 = Label(master, text='धान स्टैक(40Kg,30Kg)')
    a4.grid(row=23)
    ee3 = Entry(master)
    ee4 = Entry(master)
    ee3.grid(row=22, column=1)
    ee4.grid(row=23, column=1)
    ee3.bind("<Button-1>", some_callbackk)
    ee4.bind("<Button-1>", some_callbackk)
    ee3.insert(0,0)
    ee4.insert(0,0)
    a5 = Label(master, text='डब्लिंग')
    a5.grid(row=24)
    ee13 = Entry(master)
    ee13.bind("<Button-1>", some_callbackk)
    ee13.grid(row=24, column=1)
    ee13.insert(0,0)
    a6 = Label(master, text='चावल लोड(60kg,50kg,40kg)')
    a6.grid(row=25)
    a7 = Label(master, text='चावल लोड(25kg)')
    a7.grid(row=26)
    ee5 = Entry(master)
    ee6 = Entry(master)
    ee5.bind("<Button-1>", some_callbackk)
    ee6.bind("<Button-1>", some_callbackk)
    ee5.grid(row=25, column=1)
    ee6.grid(row=26, column=1)
    ee5.insert(0,0)
    ee6.insert(0,0)
    a8 = Label(master, text='पोलिश लोड')
    a8.grid(row=27)
    a9 = Label(master, text='चावल ढला')
    a9.grid(row=28)
    ee7 = Entry(master)
    ee8 = Entry(master)
    ee7.bind("<Button-1>", some_callbackk)
    ee8.bind("<Button-1>", some_callbackk)
    ee7.grid(row=27, column=1)
    ee8.grid(row=28, column=1)
    ee7.insert(0,0)
    ee8.insert(0,0)
    a10 = Label(master, text='बंडल स्टैक/लोड')
    a10.grid(row=29)
    a11 = Label(master, text='चावल स्टैक(25Kg)')
    a11.grid(row=30)
    ee9 = Entry(master)
    ee10 = Entry(master)
    ee9.bind("<Button-1>", some_callbackk)
    ee10.bind("<Button-1>", some_callbackk)
    ee9.grid(row=29, column=1)
    ee10.grid(row=30, column=1)
    ee9.insert(0,0)
    ee10.insert(0,0)
    a12 = Label(master, text='चावल स्टैक(50Kg,60Kg)')
    a12.grid(row=31)
    ee11 = Entry(master)
    ee11.bind("<Button-1>", some_callbackk)
    ee11.grid(row=31, column=1)
    ee11.insert(0,0)
    def export2():
        
        List2=[ee1.get(),ee2.get(),f1,ee3.get(),f2,ee4.get(),f3,ee13.get(),f11,ee5.get(),f4,ee6.get(),f5,ee7.get(),f6,ee8.get(),f7,ee9.get(),f8,ee10.get(),f9,ee11.get(),f10]
        with open('VRecords.csv', 'a',newline='') as f_object:
            df99 = pd.read_csv('VRecords.csv')

            new_df = df99.dropna()

            writer_object =csv.writer(f_object)
                    
            writer_object.writerow(List2)
            ee1.delete(0,END)
            ee2.delete(0, END)
            ee3.delete(0,END)
            ee4.delete(0,END)
            ee5.delete(0,END)
            ee6.delete(0,END)
            ee7.delete(0,END)
            ee8.delete(0,END)
            ee9.delete(0,END)
            ee10.delete(0,END)
            ee11.delete(0,END)
            ee13.delete(0,END)
            ee2.insert(0,0)
            ee3.insert(0,0)
            ee4.insert(0,0) 
            ee5.insert(0,0)
            ee6.insert(0,0)
            ee7.insert(0,0)
            ee8.insert(0,0)
            ee9.insert(0,0)
            ee10.insert(0,0)
            ee11.insert(0,0)
            ee13.insert(0,0)
    button3 = Button(master, text='Submit', width=25, command=export2)
    button3.grid(row=33,column=1)
    def Excel2():
        os.startfile("C:/Users/DELL/Downloads/S/VRecords.csv")
    button4 = Button(master, text='Excel', width=25, command=Excel2)
    button4.grid(row=33,column=2)
    def clear2():
        ee1.destroy()
        ee2.destroy()
        ee3.destroy()
        ee4.destroy()
        ee5.destroy()
        ee6.destroy()
        ee7.destroy()
        ee8.destroy()
        ee9.destroy()
        ee10.destroy()
        ee11.destroy()
        ee13.destroy()
        a1.destroy()
        a2.destroy()
        a3.destroy()
        a4.destroy()
        a5.destroy()
        a6.destroy()
        a7.destroy()
        a8.destroy()
        a9.destroy()
        a10.destroy()
        a11.destroy()
        a12.destroy()
        button3.destroy()
        button4.destroy()
        button5.destroy()
    button5 = Button(master, text='Clear', width=25, command=clear2)
    button5.grid(row=33,column=3)

def input():
    def some_callback(event):
        event.widget.delete(0, "end")
        return None
    Label(master, text='दिनांक').grid(row=0)
    Label(master, text='हौदी कटाई धान(40Kg,30Kg)').grid(row=1)
    e1 = Entry(master)
    e2 = Entry(master)
    e2.insert(0,0)
    e2.bind("<Button-1>", some_callback)
    e1.grid(row=0, column=1)
    e2.grid(row=1, column=1)
    Label(master, text='हौदी कटाई चावल(60Kg,50Kg)').grid(row=2)
    Label(master, text='धान स्टैक(40Kg,30Kg)').grid(row=3)
    e3 = Entry(master)
    e4 = Entry(master)
    e3.grid(row=2, column=1)
    e4.grid(row=3, column=1)
    e3.bind("<Button-1>", some_callback)
    e4.bind("<Button-1>", some_callback)
    e3.insert(0,0)
    e4.insert(0,0)
    Label(master, text='डब्लिंग').grid(row=4)
    e13 = Entry(master)
    e13.bind("<Button-1>", some_callback)
    e13.grid(row=4, column=1)
    e13.insert(0,0)
    Label(master, text='चावल लोड(60kg,50kg,40kg)').grid(row=5)
    Label(master, text='चावल लोड(25kg)').grid(row=6)
    e5 = Entry(master)
    e6 = Entry(master)
    e5.bind("<Button-1>", some_callback)
    e6.bind("<Button-1>", some_callback)
    e5.grid(row=5, column=1)
    e6.grid(row=6, column=1)
    e5.insert(0,0)
    e6.insert(0,0)
    Label(master, text='पोलिश लोड').grid(row=7)
    Label(master, text='चावल ढला').grid(row=8)
    e7 = Entry(master)
    e8 = Entry(master)
    e7.bind("<Button-1>", some_callback)
    e8.bind("<Button-1>", some_callback)
    e7.grid(row=7, column=1)
    e8.grid(row=8, column=1)
    e7.insert(0,0)
    e8.insert(0,0)
    Label(master, text='बंडल स्टैक/लोड').grid(row=9)
    Label(master, text='चावल स्टैक(25Kg)').grid(row=10)
    e9 = Entry(master)
    e10 = Entry(master)
    e9.bind("<Button-1>", some_callback)
    e10.bind("<Button-1>", some_callback)
    e9.grid(row=9, column=1)
    e10.grid(row=10, column=1)
    e9.insert(0,0)
    e10.insert(0,0)
    Label(master, text='चावल स्टैक(50Kg,60Kg)').grid(row=11)
    Label(master, text='एडवांस').grid(row=12)
    e11 = Entry(master)
    e12 = Entry(master)
    e11.bind("<Button-1>", some_callback)
    e12.bind("<Button-1>", some_callback)
    e11.grid(row=11, column=1)
    e12.grid(row=12, column=1)
    e11.insert(0,0)
    e12.insert(0,0)

    def export():
        
        List=[e1.get(),e2.get(),f1,e3.get(),f2,e4.get(),f3,e13.get(),f11,e5.get(),f4,e6.get(),f5,e7.get(),f6,e8.get(),f7,e9.get(),f8,e10.get(),f9,e11.get(),f10,e12.get()]
        if(e2.get()!=0 and e13.get()!=0 and e3.get()!=0 and e4.get()!=0 and e5.get()!=0 and e6.get()!=0 and e7.get()!=0 and e8.get()!=0 and e9.get()!=0 and e10.get()!=0 and e11.get()!=0 and e12!=0 ):
            with open('Records.csv', 'a',newline='') as f_object:
                df = pd.read_csv('Records.csv')

                new_df = df.dropna()

                writer_object =csv.writer(f_object)
                    
                writer_object.writerow(List)
            e1.delete(0,END)
            e2.delete(0, END)
            e3.delete(0,END)
            e4.delete(0,END)
            e5.delete(0,END)
            e6.delete(0,END)
            e7.delete(0,END)
            e8.delete(0,END)
            e9.delete(0,END)
            e10.delete(0,END)
            e11.delete(0,END)
            e12.delete(0,END)
            e13.delete(0,END)
            e2.insert(0,0)
            e3.insert(0,0)
            e4.insert(0,0) 
            e5.insert(0,0)
            e6.insert(0,0)
            e7.insert(0,0)
            e8.insert(0,0)
            e9.insert(0,0)
            e10.insert(0,0)
            e11.insert(0,0)
            e12.insert(0,0)
            e13.insert(0,0)
    button = Button(master, text='Submit', width=25, command=export)
    button.grid(row=14,column=1)

def Excel():
    os.startfile("C:/Users/DELL/Downloads/S/Records.csv")


 
while (x==True):
    input()
    button = Button(master, text='Excel', width=25, command=Excel)
    button.grid(row=18,column=1)
    button2 = Button(master, text='Vardhman', width=25, command=Vardhman)
    button2.grid(row=19,column=1)
    Sum()

    mainloop()