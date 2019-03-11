# encoding:utf-8


from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import *
import os
import xlrd

class AutoFillApplication():

    '''
    自动填写窗口 写UI的话,会让人很难过,所以这里还是用模板吧
    1.读取配置文件,获取窗口的填写栏目和变量名称
    1.1因为关键字超过20个,
    1.2.生成选项,另存为目标

    2.那么没有UI的话,就要填写模板了
    2.1首先模板通常改变的几率不大,那么我们程序中的模板就放在
    '''

    def __init__(self):
        '''
        初始化一些参数信息
        '''

        self.baseTemplet = '' #客户基础信息模板,该模板中包含了所有的key和keyname
        self.templets = [] #待替换的模板列表
        self.target = os.getcwd() + '\\new' #生成文件的存储路径,默认是执行程序下的new文件夹

        '''
        程序窗口初始化
        '''
        self.root = Tk()
        self.root.title("调查问卷上传")
        self.root.geometry('640x100')


    def _get_templets(self):
        '''
        获取相关模板信息
        1.读取模板文件列表,模板文件默认位置为执行程序下temple文件夹
        2.规则:文件命名以00开头的,为基础信息模板,确定基础信息模板
        '''
        self.templets = os.listdir(os.getcwd() + '\\temple') #读取模板文件夹下的所有模板文件名称

        for filename in self.templets:
            if(filename[:2] == '00'):
                self.baseTemplet = filename
                self.templets.remove(filename)

        return self.templets


    def _read_excel(self):
        '''
        读取baseTemplet,基础信息模板,并返回所有关键字信息
        '''
        
        allKey = [] #初始化变量列表
        
        excel_file_path = os.getcwd() + '\\temple\\' + self.baseTemplet 
        workbook = xlrd.open_workbook(excel_file_path) 
        sheet1 = workbook.sheets()[0] #读取模板  获取关键字命名,并获取关键字位置

        #遍历sheet1 获取关键字命名,并获取关键字位置
        for row in range(sheet1.nrows):
            for col in range(sheet1.ncols):
                cellText = sheet1.cell_value(row, col)
                if cellText.find('$$') > -1:
                    #找到变量后,存为字典
                    currentDic = {}
                    currentDic['key'] = cellText[2:]
                    currentDic['row'] = row
                    currentDic['col'] = col

                    allKey.append(currentDic)
        return allKey

    def _create_window(self):
        '''
        生成程序窗口
        '''

        #选择客户基本信息
        #1.如果不选择,默认是程序文件夹new中的

        #供应商选择下拉框
        name = StringVar()
        players = Combobox(self.root, textvariable=name)
        players["values"] = ("成龙", "刘德华", "周星驰")
        players.current(2)
        # players.set("演员表")
        # print(players.get())
        players.bind("<<ComboboxSelected>>", self.show_msg())
        players.pack()
        

    def show(self):
        self._create_window()
        self.root.mainloop()
    
    def show_msg(*args):
        print('呵呵')




if __name__=='__main__':

    app = AutoFillApplication()
    app.show()

    #print(app._get_templets())
    #print(app.baseTemplet)



