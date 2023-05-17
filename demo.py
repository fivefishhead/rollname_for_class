import random
import tkinter
from tkinter import *
from tkinter import ttk
import time
from tkinter import filedialog
import pandas
import windnd
from PIL import Image, ImageTk



class RandomName(ttk.Frame):
    def __init__(self, parent=None, **kw):
        ttk.Frame.__init__(self, parent, **kw)
        self.name_list = ['请上传名单']
        self._start = 0.0
        self._elapsedtime = 0.0
        self._running = False
        self.timestr = StringVar()
        self.style = ttk.Style(self)
        self.tk.call("source", "forest-light.tcl")
        self.style.theme_use("forest-light")
        self.var_count = IntVar()
        self.var_count.set(0)
        self.strs = []
        for i in range(0, 9):
            name_str = StringVar()
            name_str.set(' ')
            self.strs.append(name_str)
        self.name_count = 0
        self.name_temp = self.name_list[0]
        self.makeWidgets()
        self.drag()

    def drag(self):
        windnd.hook_dropfiles(self, func=self.drag_files)
        return


    def drag_files(self, urls):
        path = urls[0].decode("gbk")
        if path.endswith('xlsx') or path.endswith('xls'):
            excel_file = pandas.read_excel(path, header=None)
            for j in range(len(excel_file.index)):
                for i in range(len(excel_file.loc[j].values)):
                    if(excel_file.loc[j].values[i] =='姓名'):
                        #print('成功获取姓名列表\n')
                        df = excel_file.iloc[j+1:, i].dropna(axis=0, how='all')
                        self.name_list = df.values.tolist()
                        self.var_count.set(len(self.name_list))
                        #print(len(self.name_list))


    def makeWidgets(self):
        #  定义标签栏
        l = ttk.Label(self, textvariable=self.timestr, font=("Arial, 30"))
        l.pack(side=TOP)
        self.set_name(self._elapsedtime)
        count = ttk.Label(self, text='总人数', font=("Arial, 10"))
        count.pack(side=TOP)
        name_count = ttk.Label(self, textvariable=self.var_count, font=("Arial, 10"))
        name_count.pack(side=TOP)
        for i in range(0, 9):
            temp = ttk.Label(self, textvariable=self.strs[i], font=("Arial, 25"))
            temp.pack(side=TOP)


    def update(self):
        # 更新显示内容
        self._elapsedtime = time.time() - self._start
        self.set_name(self._elapsedtime)  # 设置显示内容
        self._timer = self.after(50, self.update)  # 刷新界面

    def set_name(self, elap):
        # 随机产生姓名
        cur = int(elap * 100) % (len(self.name_list))
        self.timestr.set(self.name_list[cur])
        self.name_temp = self.name_list[cur]


    def upload_file(self):
        path = tkinter.filedialog.askopenfilename()  # askopenfilename 1次上传1个；askopenfilenames1次上传多个
        if path.endswith('xlsx') or path.endswith('xls'):
            excel_file = pandas.read_excel(path, header=None)
            for j in range(len(excel_file.index)):
                for i in range(len(excel_file.loc[j].values)):
                    if(excel_file.loc[j].values[i] =='姓名'):
                        df = excel_file.iloc[j+1:, i].dropna(axis=0, how='all')
                        self.name_list = df.values.tolist()
                        self.var_count.set(len(self.name_list))
                        #print(len(self.name_list))



    def Start(self):
        # 开始
        if not self._running:
            self._start = time.time() - self._elapsedtime
            self.update()
            self._running = True

    def Stop(self):
        # 暂停
        if self._running:
            self.after_cancel(self._timer)
            self._elapsedtime = time.time() - self._start
            self.set_name(self._elapsedtime)
            self._running = False
            self.strs[self.name_count % 9].set(self.name_temp)
            self.name_count += 1
            random.shuffle(self.name_list)

    def clear(self):
        for i in range(0, 9):
            self.strs[i].set('')
        self.name_count = 0

    def name_label(self):
        # 显示窗口
        self.pack(side=TOP)
        ttk.Button(self, text='开始', command=self.Start, width=6, style='Accent.TButton').pack(side=LEFT)
        ttk.Button(self, text='暂停', command=self.Stop, width=6, style='Accent.TButton').pack(side=LEFT)
        ttk.Button(self, text='上传', command=self.upload_file, width=6, style='Accent.TButton').pack(side=LEFT)
        ttk.Button(self, text='清空', command=self.clear, width=6, style='Accent.TButton').pack(side=LEFT)



if __name__ == '__main__':

    root = tkinter.Tk()
    root.title('马克思选中的孩子')
    root.geometry('800x500+300+300')
    root.resizable(False, False)
    canvas_root = tkinter.Canvas(root, width=800, height=500)
    im_root = ImageTk.PhotoImage(Image.open(r'./background.jpeg').resize((500, 500)))
    canvas_root.create_image(250, 250, image=im_root)
    canvas_root.pack()
    message_frame = tkinter.Frame(root)
    message_frame.configure(bd=2)
    message_frame.place(x=510, y=30, width=280, height=460)
    sw = RandomName(message_frame)
    sw.name_label()
    root.mainloop()