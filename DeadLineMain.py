import tkinter  as tk
from tkinter import messagebox as mb
import tkinter.font as tkFont
import easygui
import pandas as pd
import re
class DeadLine:

    def __init__(self):
        self.monthDict = {1: 31, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}



        self.run()

    def run(self):
        self.clickCount=0
        self.win=tk.Tk()
        self.win.title("期限判斷器")
        self.win.geometry('950x450+500+400')
        f = open("預設年份設定檔.txt", mode='r')
        fontStyle = tkFont.Font(family="Lucida Grande", size=15)
        self.label_year=tk.Label(self.win,text="年份設定(民國):",font=fontStyle).grid(row=0,column=0)

        self.entry_year = tk.Entry(self.win, width=10, font=fontStyle)
        self.entry_year.insert(tk.END, f.read(3))
        self.entry_year.grid(row=0, column=1)

        self.RC_year = str(self.entry_year.get())
        self.AD_Year = self.ConvertYear(int(self.RC_year))
        self.monthDict[2] = self.getYearDays(self.AD_Year)




        self.label_input=tk.Label(self.win,text="輸入檔案路徑:",font=fontStyle).grid(row=1,column=0)
        self.btn_input=tk.Button( self.win,text='瀏覽檔案(輸入)',command=self.btn_input_getFileName,font=fontStyle).grid(row=1,column=1)


        self.label_input=tk.Label(self.win,text="輸出檔案路徑:",font=fontStyle).grid(row=2,column=0)
        self.entry_input = tk.Entry(self.win,width=70,font=fontStyle)
        self.entry_input.grid(row=1,column=2)

        self.entry_output = tk.Entry(self.win, width=70,font=fontStyle)
        self.entry_output.grid(row=2,column=2)


        self.btn_start=tk.Button( self.win,text='開始轉換',command=self.btn_begin_convert,font=fontStyle).grid(row=3,column=0)

        scrollbar = tk.Scrollbar(self.win)
        scrollbar.grid(row=3, column=1, columnspan=2, rowspan=1)
        self.showText=tk.Text(self.win,yscrollcommand=scrollbar.set)
        self.showText.grid(row=3, column=1, columnspan=2, rowspan=1)

        self.label_AWen = tk.Label(self.win, text="DeadLine Program Make by AWen 2021/3/28 v1.1", font=tkFont.Font(family="Lucida Grande", size=9)).grid(row=5, column=2,columnspan=2)
        self.btn_fix = tk.Button(self.win, text='更新紀錄', command=self.btn_updateList, font=fontStyle).grid(
            row=5, column=0)
        self.win.mainloop()

    def btn_updateList(self):
        mb.showinfo("更新紀錄",
            """-----------------------\n更新日期:2021/4/5\n更新版本:v1.1\n更新內容:修正輸入月份格式問題\n新增PermissionError問題\n-----------------------
            """)

    def btn_input_getFileName(self):


        self.file_path = easygui.fileopenbox()
        print(self.file_path)
        self.entry_input.delete(0,tk.END)
        self.entry_input.insert(0,self.file_path)
        self.entry_output.delete(0,tk.END)
        temp=self.file_path.split('.')
        outputName=temp[0]+"_已轉檔."+temp[1]
        self.entry_output.insert(0, outputName)
        # self.text_input.insert(0,file_path)

    def btn_begin_convert(self):
        if self.clickCount != 0:
            self.showText.delete('1.0','end')

        self.ansList = []
        self.RC_year = str(self.entry_year.get())
        self.AD_Year = self.ConvertYear(int(self.RC_year))
        self.monthDict[2] = self.getYearDays(self.AD_Year)
        self.inputPath=str(self.entry_input.get())
        try:

            df = pd.read_excel(self.inputPath)
            # excel = pd.ExcelFile("test.xlsx")
            # print(excel.sheet_names)
            colName=df.columns

            for index, row in df.iterrows():
                m=re.search('[1-12]*',str(row[colName[0]]))
                self.getDeadLine(int(m.group(0)), row[colName[1]])

            mb.showinfo("確認", "已轉換", detail="請於\"輸出檔案路徑\"確認檔案")
            print(df)
            self.clickCount += 1
        except AssertionError :
            mb.showerror("錯誤!", "輸入框錯誤", detail="請確認\"輸入檔案路徑\"是否正確")
        except PermissionError:
            mb.showerror("錯誤!", "請先關閉輸出檔案!", detail="關閉所有excel檔案重試一次")
    def getDeadLine(self,month,days):


        ans=""
        self.ansYear=self.AD_Year
        print('題目:{0}年 {1} 月 {2} 日'.format(int(self.AD_Year),int(month),int(days)))
        # print(self.monthDict[int(month)])
        if days%30==0:
            mCount = days // 30
            if 30 > days and self.getDeadMonth(month,mCount) == 2:
                ans = '{0} 月 {1} 日'.format(int(self.getDeadMonth(month, mCount)),int(self.getYearDays(self.AD_Year)))
            # ans=str(month)+"月"+str(self.monthDict[month])+'日'
            else:
                tempM=self.getDeadMonth(month,mCount)
                ans='{0} 月 {1} 日'.format(int(tempM),int(self.monthDict[tempM]))
        # elif month==12:
        #     mCount = days // 30
        #
        #     # ans=str(month)+"月"+str(self.monthDict[month])+'日'
        #     ans = '{0} 月 {1} 日'.format(self.getDeadMonth(month, mCount),
        #                                self.monthDict[self.getDeadMonth(month, mCount)])

        elif days < 30:
            ans = '{0} 月 {1} 日'.format(int(self.getDeadMonth(month,1)),int(days))
        elif days > 30:
            mCount=days//30+1
            daysTemp=days%30
            if month==12:
                ans = '{0} 月 {1} 日'.format(int(self.getDeadMonth(month,mCount)), int(daysTemp))

            else:
                ans = '{0} 月 {1} 日'.format(int(self.getDeadMonth(month,mCount)), int(daysTemp))
            print(mCount)
            print(daysTemp)
        self.ansList .append([str(month), days,self.ansYear-1911,ans ])
        print(str(self.ansYear)+"年 "+ans)
        self.showText.insert(tk.END, str(int(month))+"月"+str(int(days))+"日 => "+str(self.ansYear)+"年 "+ans+'\n')
        ansDF = pd.DataFrame(self.ansList ,
                   columns=['月份', '天數','兌現日期(年)','兌現日期'])
        ansDF.to_excel(str(self.entry_output.get()))


    def getDeadMonth(self,m,count):
        if m==12 and count==0:
            return m
        elif m+count>12:
            self.ansYear+=1
            return (m+count)%12
        else:
            return (m+count)%12


    def getYearDays(self,year):
        days=-1
        if (year % 4) == 0:
            if (year % 100) == 0:
                if (year % 400) == 0:
                    days=29
                    print("{0} 是闰年".format(year))  # 整百年能被400整除的是闰年
                else:
                    days = 28
                    print("{0} 不是闰年".format(year))
            else:
                days = 29
                print("{0} 是闰年".format(year))  # 非整百年能被4整除的为闰年
        else:
            days = 28
            print("{0} 不是闰年".format(year))
        return days
# win.resizable(False, False)

    def ConvertYear(self,year):
        return year+1911


DeadLine()


