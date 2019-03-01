# encoding:utf-8

import docx
import copy
import  os
import xlrd



'''
allKey = [

    dic = {
    
        key : ...  关键字名称
        row : ...  模板中的行数
        col : ...  模板中的列数
        value : ...  新客户的信息
    }
]

'''
        
#替换关键字
def fill_file(allKey, filename):
    if filename.find(".docx") > -1:
        filepath = os.getcwd() + '\\temple\\' + filename
        doc = docx.Document(filepath)
        for para in doc.paragraphs:
            #runs = copy.deepcopy(para.runs)
            for run in para.runs:

                #进行替换关键字
                for dic in allKey:
                    if run.text.find(dic['key']) > -1:
                        run.text = run.text.replace(dic['key'], dic['value'])
                        #para.clear()
                        #para.add_run(run.text, run.style)
        newfilepath = os.getcwd() + '\\new\\' + filename.replace('docx', 'doc')
        doc.save(newfilepath)
    elif filename.find('.xls') > -1:
        print('是个表')
    else:
        print('不知道是个啥')

#获取关键字
def read_excel():

    #初始化变量列表
    allKey = []

    ''' 读取信息采集文件,获取关键字 '''
    #首先读取模板  获取关键字命名,并获取关键字位置
    excelName = os.getcwd() + '\\temple\\00客户信息采集.xls'

    workbook = xlrd.open_workbook(excelName)
    sheet1 = workbook.sheets()[0]
    #遍历sheet1
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


def read_newCustomerExcel(allKey):
    excelName = os.getcwd() + '\\new\\00客户信息采集.xls'
    workbook = xlrd.open_workbook(excelName)
    sheet1 = workbook.sheets()[0]

    #遍历所有的dic
    for dic in allKey:
        #获取新客户信息
        ctype = sheet1.cell(dic['row'], dic['col']).ctype 
        if ctype == 2 and ctype % 1 == 0.0:  # ctype为2且为浮点
            dic['value'] = str(int(sheet1.cell_value(dic['row'], dic['col'])))
        else :
            dic['value'] = str(sheet1.cell_value(dic['row'], dic['col']))
    return allKey

#获取模板列表
def templeFiles():
    return os.listdir(os.getcwd() + '\\temple')

#进行替换
def autoFill():
    
    allKey = read_excel()
    #1.根据关键字获取新客户的信息
    allKey = read_newCustomerExcel(allKey)
    
    #2.将获取到的客户信息,自动填充到各模板并生成新的文件
    for file in templeFiles():
        fill_file(allKey, file)

    
                  
if __name__ == "__main__":
    #print(templeFiles())
    autoFill()
    #read_docx('你好你好.docx');
#read_docx("C:\\Users\\Administrator\\Desktop\\python_word\\你好你好.docx")