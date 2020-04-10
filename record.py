from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk  #不报错
from xlutils.copy import copy
import xlwt 
import xlrd
from os import remove

class Base():
    def __init__(self, master):
        self.root = master
        self.root.config()
        self.root.title("咨询信息记录查询系统")
        self.root.geometry("1000x680")
        self.root.resizable(False, False)
        Mainface(self.root)

class Mainface():
    def __init__(self, master):
        self.master = master
        self.master.config(bg="palegoldenrod")
        #生成图片
        self.Pilimage = Image.open(r"Image\background.gif")  #指定为self,不会报错
        self.image = ImageTk.PhotoImage(image=self.Pilimage)
        self.mainface = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.mainface.pack()
        b_record = Button(self.mainface, text="新 建", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.record) #按钮1
        b_record.place(relx=0.3, rely=0.3, anchor=CENTER)
        b_search = Button(self.mainface, text="查 询", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue",
command=self.search) #按钮2
        b_search.place(relx=0.7, rely=0.3, anchor=CENTER)
        l_image = Label(self.mainface, image=self.image)
        l_image.place(relx=0.5, rely=0.5, anchor=N)
        
    def record(self):
        self.mainface.destroy()
        Record(self.master)
        
    def search(self):
        self.mainface.destroy()
        Search(self.master)
            
class Record():
    def __init__(self, master):
        self.master = master
        self.master.config(bg="palegoldenrod")
        self.record = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.record.pack()
        #设置列宽
        self.record.grid_columnconfigure(0, minsize=100)
        self.record.grid_columnconfigure(1, minsize=200)
        self.record.grid_columnconfigure(2, minsize=200)
        self.record.grid_columnconfigure(3, minsize=200)
        self.record.grid_columnconfigure(4, minsize=200)
        self.record.grid_columnconfigure(5, minsize=100)
        self.record.grid_rowconfigure(8, minsize=100)
        #第一行
        #姓名
        self.var_name = StringVar()
        self.var_name.set("编号或姓名")
        self.l_name = Label(self.record, text="姓 名：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=0, pady=5, sticky=E)
        self.e_name = Entry(self.record, textvariable=self.var_name,font=(r"Font\simhei.ttf", 12))
        self.e_name.grid(row=0, column=1, pady=5, sticky=W)
        
        #咨询日期
        self.l_date = Label(self.record, text="咨询日期（xxxx-xx-xx）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=2, pady=5, sticky=E)
        self.e_date = Entry(self.record, font=(r"Font\simhei.ttf", 12))
        self.e_date.grid(row=0, column=3, pady=5, sticky=W)
        
        #第二行
        #使用物质
        self.l_substance = Label(self.record, text="使用物质：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=0, pady=5, sticky=E)
        self.var_subs = StringVar()
        self.var_subs.set("冰毒")
        self.O_substance = OptionMenu(self.record, self.var_subs, "冰毒","海洛因","大麻","合成大麻","其他").grid(row=1, column=1, pady=5, sticky=W)
        #是否尿检
        self.l_un = Label(self.record, text="是否尿检及结果：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=2, pady=5, sticky=E)
        self.var_un = StringVar()
        self.var_un.set("否")
        self.O_substance = OptionMenu(self.record, self.var_un, "否","是").grid(row=1, column=3, pady=5, sticky=W)
        self.var_un_result = StringVar()
        self.var_un_result.set("-")
        self.O_substance = OptionMenu(self.record, self.var_un_result, "-","阴性","阳性").grid(row=1, column=4, pady=5, sticky=W)
        
        #第三行
        #年龄
        self.l_age = Label(self.record, text="年龄（岁）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=0, pady=5, sticky=E)
        self.e_age = Entry(self.record, font=(r"Font\simhei.ttf", 12))
        self.e_age.grid(row=2, column=1, pady=5, sticky=W)
        #精神症状
        self.l_mental = Label(self.record, text="是否有精神症状：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=2, pady=5, sticky=E)
        self.var_mental = StringVar()
        self.var_mental.set("否")
        self.O_substance = OptionMenu(self.record, self.var_mental, "否","是").grid(row=2, column=3, pady=5, sticky=W)
        
        #第四、五行
        #诊断
        self.l_diagnose = Label(self.record, text="病情诊断：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=0, pady=5, sticky=E)
        self.e_diagnose = Text(self.record, height=3, font=(r"Font\simhei.ttf", 12))
        self.e_diagnose.grid(row=3, column=1, columnspan=4, ipady=12, pady=5, sticky=W)
        
        #第六、七、八行
        #咨询内容记录
        self.l_content = Label(self.record, text="咨询内容记录：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=5, column=0, pady=5, sticky=E)
        self.e_content = Text(self.record, height=15, font=(r"Font\simhei.ttf", 12))
        self.e_content.grid(row=5, column=1, columnspan=4, ipady=36, pady=5, sticky=W)
        
        #文件保存
        self.b_save = Button(self.record, text="保 存", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.save)
        self.b_save.grid(row=8, column=0, columnspan=3, pady=5, sticky=S)
        #返回键
        self.record_back = Button(self.record, text="返 回", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.back)
        self.record_back.grid(row=8, column=2, columnspan=3, pady=5, sticky=S)
    
    def back(self):
        self.record.destroy()
        Mainface(self.master)
    
    #文件保存
    def save(self):
        ftypes = [("Excel files", ".xls"), ("All files", " *")] 
        file_name = self.e_name.get()
        file_date = self.e_date.get()
        file_subs = self.var_subs.get()
        file_un = self.var_un.get()
        file_un_result = self.var_un_result.get()
        file_age = self.e_age.get()
        file_mental = self.var_mental.get()
        file_diagnose = self.e_diagnose.get(1.0, END)
        file_content = self.e_content.get(1.0, END)
        #file_path = filedialog.asksaveasfilename(title="保存文件", filetypes=ftypes, defaultextension=".xls")
        file_path = "DataFile\\" + file_name + ".xls"
        print(file_name)
        messagebox.showinfo("提示", "保存成功！")
        if file_path is not None:
            book = xlwt.Workbook(encoding="utf-8", style_compression=0)
            sheet = book.add_sheet(file_name, cell_overwrite_ok=True)
            sheet.write(0, 0, "姓 名")
            sheet.write(1, 0, file_name)
            sheet.write(0, 1, "咨询日期")
            sheet.write(1, 1, file_date)
            sheet.write(0, 2, "使用物质")
            sheet.write(1, 2, file_subs)
            sheet.write(0, 3, "是否尿检")
            sheet.write(1, 3, file_un)
            sheet.write(0, 4, "尿检结果")
            sheet.write(1, 4, file_un_result)
            sheet.write(0, 5, "年 龄")
            sheet.write(1, 5, file_age)
            sheet.write(0, 6, "是否有精神症状")
            sheet.write(1, 6, file_mental)
            sheet.write(0, 7, "病情诊断")
            sheet.write(1, 7, file_diagnose)
            sheet.write(0, 8, "咨询内容记录")
            sheet.write(1, 8, file_content)
            #保存
            book.save(file_path)
        
class Search():
    def __init__(self, master):
        self.master = master
        self.master.config(bg="palegoldenrod")
        self.search = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.search.pack()
        #设置列宽
        self.search.grid_columnconfigure(0, minsize=100)
        self.search.grid_columnconfigure(1, minsize=200)
        self.search.grid_columnconfigure(2, minsize=200)
        self.search.grid_columnconfigure(3, minsize=200)
        self.search.grid_columnconfigure(4, minsize=200)
        self.search.grid_columnconfigure(5, minsize=100)
        self.search.grid_rowconfigure(9, minsize=150)
        #设置变量
        self.name = StringVar()
        self.name.set("")
        self.date = StringVar()
        self.date.set("-")
        self.substance = StringVar()
        self.substance.set("")
        self.un = StringVar()
        self.un.set("")
        self.un_result = StringVar()
        self.un_result.set("")
        self.age = StringVar()
        self.age.set("")
        self.mental = StringVar()
        self.mental.set("")
        self.times = StringVar()
        self.times.set("")
        self.diagnose = ""
        self.content = ""
        
        #第一行
        #查询
        self.l_search = Label(self.search, text="请输入编号或姓名：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=0, columnspan=2, pady=5, sticky=E)
        self.e_search = Entry(self.search, font=(r"Font\simhei.ttf", 12))
        self.e_search.grid(row=0, column=2, columnspan=2, ipadx=10, pady=5, sticky=W)
        self.b_search = Button(self.search, text="查 询", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue",
command=self.searchMethod).grid(row=0, column=4, pady=5, sticky=W)
        
        #第二行
        #姓名
        self.l_name1 = Label(self.search, text="姓 名：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=0, pady=15, sticky=E)
        self.l_name2 = Label(self.search, textvariable=self.name, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=1, pady=5, sticky=W)
        #咨询日期
        self.l_date1 = Label(self.search, text="咨询日期（xxxx-xx-xx）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=2, pady=5, sticky=E)
        self.l_date2 = Label(self.search, textvariable=self.date, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=3, pady=5, sticky=W)
        
        #第三行
        #物质使用
        self.l_substance1 = Label(self.search, text="使用物质：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=0, pady=5, sticky=E)
        self.l_substance2 = Label(self.search, textvariable=self.substance, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=1, pady=5, sticky=W)
        #是否尿检及其结果
        self.l_un1 = Label(self.search, text="是否尿检及结果：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=2, pady=5, sticky=E)
        self.l_un2 = Label(self.search, textvariable=self.un, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=3, pady=5, sticky=W)
        self.l_un_result1 = Label(self.search, textvariable=self.un_result, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=4, pady=5, sticky=W)
        
        #第四行
        #年龄
        self.l_age1 = Label(self.search, text="年龄（岁）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=0, pady=5, sticky=E)
        self.l_age2 = Label(self.search, textvariable=self.age, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=1, pady=5, sticky=W)
        #精神症状
        self.l_mental1 = Label(self.search, text="是否有精神症状：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=2, pady=5, sticky=E)
        self.l_mental2 = Label(self.search, textvariable=self.mental, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=3, pady=5, sticky=W)
        #来访次数
        self.l_times1 = Label(self.search, text="来访次数：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=3, pady=5, sticky=E)
        self.l_times2 = Label(self.search, textvariable=self.times, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=4, pady=5, sticky=W)
        
        #第五行
        #诊断
        self.l_diagnose1 = Label(self.search, text="病情诊断：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=4, column=0, pady=5, sticky=E)
        self.l_diagnose2 = Text(self.search, height=3, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod")
        self.l_diagnose2.grid(row=4, column=1, columnspan=4, pady=5, sticky=W)
        self.l_diagnose2.insert(END, self.diagnose) 
        
        #第六行
        #咨询内容
        self.l_content1 = Label(self.search, text="咨询内容记录：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=5, column=0, pady=5, sticky=E)
        self.l_content2 = Text(self.search, height=10, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod")
        self.l_content2.grid(row=5, column=1, columnspan=5, pady=5, sticky=W)
        self.l_content2.insert(END, self.content)
        
        #添加按钮
        search_add = Button(self.search, text="添 加", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue",
command=self.add)
        search_add.grid(row=9, column=2, columnspan=2, pady=5, sticky=W)
        
        #返回按钮
        search_back = Button(self.search, text="返 回", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue",
command=self.back)
        search_back.grid(row=9, column=3, columnspan=2, pady=5, sticky=W)
    
    #查询方法
    def searchMethod(self):
        self.search_path = "DataFile\\" + str(self.e_search.get()) + ".xls"
        try:
            xl = xlrd.open_workbook(self.search_path)
            table = xl.sheets()[0]
            self.name.set(table.cell(1, 0).value)
            self.substance.set(table.cell(1, 2).value)
            self.un.set(table.cell(1, 3).value)
            self.un_result.set(table.cell(1, 4).value)
            self.age.set(table.cell(1, 5).value)
            self.mental.set(table.cell(1, 6).value)
            self.diagnose += str(table.cell(1, 7).value)
            self.l_diagnose2.insert(END, self.diagnose)
            num = 0
            for i in range(100):
                try:
                    varBlank = table.cell(i, 8).value
                except:
                    num = i
                    break
            self.times.set(str(num-1))
            for ii in range(1, num):
                self.content = self.content + str(table.cell(ii, 1).value) + "\n" + str(table.cell(ii, 8).value) + "\n\n"
                       
            self.l_content2.insert(END, self.content)
                    
        except:
            messagebox.showinfo("提示", "没有该患者！")
            self.search.destroy()
            Search(self.master)
      
    #添加方法
    def add(self):
        try:
            self.search.destroy()
            Add(self.master, self.search_path)
        except:
            messagebox.showinfo("提示", "请先进行查询！")
            self.search.destroy()
            Search(self.master)
        
    #返回方法   
    def back(self):
        self.search.destroy()
        Mainface(self.master)  
        
#添加界面
class Add():
    def __init__(self, master, path):
        self.master = master
        self.path = path
        self.master.config(bg="palegoldenrod")
        self.add = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.add.pack()
        #设置列宽
        self.add.grid_columnconfigure(0, minsize=200)
        self.add.grid_columnconfigure(1, minsize=200)
        self.add.grid_columnconfigure(2, minsize=200)
        self.add.grid_columnconfigure(3, minsize=200)
        self.add.grid_columnconfigure(4, minsize=100)
        self.add.grid_columnconfigure(5, minsize=100)
        self.add.grid_rowconfigure(3, minsize=150)
        
        #第一行  咨询日期
        self.l_date = Label(self.add, text="咨询日期（xxxx-xx-xx）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=0, pady=5, sticky=E)
        self.e_date = Entry(self.add, font=(r"Font\simhei.ttf", 12))
        self.e_date.grid(row=0, column=1, pady=15, sticky=W)
        #第二行  咨询内容
        self.l_content = Label(self.add, text="咨询内容记录：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=0, pady=5, sticky=E)
        self.e_content = Text(self.add, height=15, font=(r"Font\simhei.ttf", 12))
        self.e_content.grid(row=1, column=1, columnspan=4, ipady=36, pady=5, sticky=W)
        
        #确认添加按钮
        self.b_confirmAdd = Button(self.add, text="保 存", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.confirmAdd)
        self.b_confirmAdd.grid(row=3, column=1, pady=5, sticky=S)
        #退出按钮
        self.b_quit = Button(self.add, text="返 回", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.back)
        self.b_quit.grid(row=3, column=3, pady=5, sticky=S)
     
    #确认添加方法
    def confirmAdd(self):
        date_a = StringVar()
        xl_a = xlrd.open_workbook(self.path)
        table = xl_a.sheets()[0]
        table_a = copy(xl_a)
        sheet = table_a.get_sheet(0)
        add_date = self.e_date.get()
        add_content = self.e_content.get(1.0, END)
        num_a = 0
        for i in range(100):
            try:
                varBlank_a = table.cell(i, 8).value
            except:
                num_a = i
                break
        sheet.write(num_a, 1, add_date)
        sheet.write(num_a, 8, add_content)
        remove(self.path)
        table_a.save(self.path)
        messagebox.showinfo("提示", "保存成功！")
            
            
    #返回键    
    def back(self):
        self.add.destroy()
        Mainface(self.master)
        
        
if __name__ == "__main__":
    try:
        root = Tk()
        Base(root)
        root.mainloop()
    except SystemExit as msg:
        print(msg)
       



