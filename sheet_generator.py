import pandas as pd
import openpyxl
import os
import xlrd,xlsxwriter

def sheet_generator(file_name, sheet_name):
    '''
    生成表格
    :param file_name:
    :param sheet_name:
    :return:
    '''
    while True:
        # 读取原始表
        df = pd.read_excel('u-fold/原始表.xlsx',sheet_name=sheet_name)
        # 随机排序
        df = df.sample(frac=1).reset_index(drop=True).drop(['序号'], 1)
        df.index = df.index + 1
        # 添加同时检测标记
        df, flag = together(df = df, sheet_name=sheet_name, times=3)
        if flag is False:
            continue
        df, flag = together(df=df, sheet_name=sheet_name, times=2)
        if flag is False:
            continue
        else:
            break
    df.index.name = '序号'
    # 储存
    save(file_name,file_name + sheet_name,df)

def seperater(file_name, sheet_name):
    '''
    将表格分开存储，方便打印成纸质文件
    :param df:
    :return:
    '''
    # 分次储存
    df = pd.read_excel(file_name + '.xlsx', sheet_name=file_name + sheet_name,index_col=0)
    df_list = []
    if sheet_name == '夏装':
        df_list.append(df.iloc[0:24])
        df_list.append(df.iloc[24:48])
        df_list.append(df.iloc[48:56])
    elif sheet_name == '冬装' or sheet_name == '春秋':
        df_list.append(df.iloc[0:24])
        df_list.append(df.iloc[24:48])
        df_list.append(df.iloc[48:72])
    else:
        print('sheet_name 必须为 冬装、春秋或夏装')
    #
    if not os.path.exists('sheet'):
        os.mkdir('sheet')


    for i,df0 in enumerate(df_list):
        save('sheet/' + file_name, file_name + sheet_name + str(i),df0)

def word(file_name, sheet_name):
        for i in range(3):
            l1 = get_file_value('u-fold/表头/' + file_name + '.xlsx', file_name + sheet_name)
            fn = sheet_name + str(i)
            l2 = get_file_value1('sheet/'+ file_name + '.xlsx', fn )
            for l in l2:
                l1.append(l)
            data_write('temp','1',l1)
            df = pd.read_excel('temp.xlsx',index_col=0)
            save('sheet/word' + file_name,file_name + sheet_name + str(i),df)

def together(df, times, sheet_name):
    i = 1
    num = 0
    while True:
        if i >= df.shape[0] - times:
            print('没有了%d' % i)
            return df, False
        df_sub = df[i:i + times]
        if 1. in df_sub['同时测试'].values.tolist():
            i = i + 1
            continue
        if 2. in df_sub['同时测试'].values.tolist():
            i = i + 1
            continue
        if 3. in df_sub['同时测试'].values.tolist():
            i = i + 1
            continue
        print(sheet_name,times)
        plist = df_sub['违禁物品或其模拟物代号'].values.tolist()
        qlist = df_sub['身体位置代号'].values.tolist()

        plist_d = []
        if plist[0] == 1:
            plist_d.append(1)
        elif plist[0] == 2 or plist[0] == 3 or plist[0] == 4:
            plist_d.append(2)
        elif plist[0] == 5 or plist[0] == 6:
            plist_d.append(3)
        elif plist[0] == 7 or plist[0] == 8:
            plist_d.append(4)
        else:
            i = i + 1
            continue

        if plist[1] == 1:
            plist_d.append(1)
        elif plist[1] == 2 or plist[1] == 3 or plist[1] == 4:
            plist_d.append(2)
        elif plist[1] == 5 or plist[1] == 6:
            plist_d.append(3)
        elif plist[1] == 7 or plist[1] == 8:
            plist_d.append(4)
        else:
            i = i + 1
            continue
        if times == 3:
            if plist[2] == 1:
                plist_d.append(1)
            elif plist[2] == 2 or plist[2] == 3 or plist[2] == 4:
                plist_d.append(2)
            elif plist[2] == 5 or plist[2] == 6:
                plist_d.append(3)
            elif plist[2] == 7 or plist[2] == 8:
                plist_d.append(4)
            else:
                i = i + 1
                continue

        qlist_d = []
        for q in qlist:
            if q == 'A':
                qlist_d.append(1)
            elif q == 'B':
                qlist_d.append(2)
            elif q == 'C':
                qlist_d.append(3)
            elif q == 'D':
                qlist_d.append(4)
            elif q == 'E':
                qlist_d.append(5)
            elif q == 'F':
                qlist_d.append(6)
            elif q == 'G':
                qlist_d.append(7)
            elif q == 'H':
                qlist_d.append(8)

        if len(plist_d) == len(set(plist_d)):
            if len(qlist_d) == len(set(qlist_d)):
                print('---%d'%i)
                print(qlist)
                print(qlist_d)
                if times == 3:
                    fl = abs(list(set(qlist_d))[1] -  list(set(qlist_d))[0]) > 1 and abs(list(set(qlist_d))[1] -  list(set(qlist_d))[2]) > 1 and abs(list(set(qlist_d))[2] -  list(set(qlist_d))[0]) > 1
                elif times == 2:
                    fl = abs(list(set(qlist_d))[1] -  list(set(qlist_d))[0]) > 1
                if fl:
                    num = num + 1
                    if times ==3:
                        df.loc[i+1:i + times]['同时测试'] = [3., 3., 3.]
                    elif times == 2:
                        df.loc[i+1:i + times]['同时测试'] = [2., 2.]
                    print(df.loc[i:i + times]['同时测试'])
                    i = i + times + 1
                    if sheet_name == '夏装':
                        if num == 2:
                            break
                    if num == 3:
                        break
                    continue
        i = i + 1
    return df ,True

def open_xls(file):
    try:
        fh=xlrd.open_workbook(file)
        return fh
    except Exception as e:
        print("打开文件错误："+e)

#根据excel名以及第几个标签信息就可以得到具体标签的内容
def get_file_value(filename,sheetname):
    if sheetname == filename + '冬装':
        sheetnum = 0
    elif sheetname == filename + '春秋':
        sheetnum = 1
    else:
        sheetnum = 2
    rvalue=[]
    fh=open_xls(filename)
    sheet=fh.sheets()[sheetnum]
    row_num=sheet.nrows
    for rownum in range(0,row_num):
        rvalue.append(sheet.row_values(rownum))
    return rvalue

def get_file_value1(filename,sheetname):
    if sheetname == '冬装0':
        sheetnum = 0
    elif sheetname ==  '冬装1':
        sheetnum = 1
    elif sheetname ==  '冬装2':
        sheetnum = 2
    elif sheetname ==  '春秋0':
        sheetnum = 3
    elif sheetname ==  '春秋1':
        sheetnum = 4
    elif sheetname ==  '春秋2':
        sheetnum = 5
    elif sheetname ==  '夏装0':
        sheetnum = 6
    elif sheetname ==  '夏装1':
        sheetnum = 7
    elif sheetname ==  '夏装2':
        sheetnum = 8
    rvalue=[]
    fh=open_xls(filename)
    sheet=fh.sheets()[sheetnum]
    row_num=sheet.nrows
    for rownum in range(0,row_num):
        rvalue.append(sheet.row_values(rownum))
    return rvalue


#  将数据写入新文件
def data_write(file_path,sheetname, datas):
    f = xlsxwriter.Workbook(file_path + '.xlsx')
    sheet1 = f.add_worksheet(sheetname)

    # 将数据写入第 i 行，第 j 列
    i = 0
    for data in datas:
        for j in range(len(data)):
            sheet1.write(i, j, data[j])
        i = i + 1

    f.close()  # 保存文件

    # 储存
def save(file_name, sheet_name, df):
    if  not os.path.exists(file_name + '.xlsx'):
        df.to_excel(file_name + '.xlsx', sheet_name= sheet_name)
    else:
        excelWriter = pd.ExcelWriter(file_name + '.xlsx',engine='openpyxl')
        book = openpyxl.load_workbook(excelWriter.path)
        excelWriter.book = book
        df.to_excel(excel_writer=excelWriter, sheet_name=sheet_name)
        excelWriter.close()

if __name__ == '__main__':
    sheet_generator('男1', '冬装')
    sheet_generator('男1', '春秋')
    sheet_generator('男1', '夏装')
    sheet_generator('男2', '冬装')
    sheet_generator('男2', '春秋')
    sheet_generator('男2', '夏装')
    sheet_generator('男3', '冬装')
    sheet_generator('男3', '春秋')
    sheet_generator('男3', '夏装')
    sheet_generator('女1', '冬装')
    sheet_generator('女1', '春秋')
    sheet_generator('女1', '夏装')
    sheet_generator('女2', '冬装')
    sheet_generator('女2', '春秋')
    sheet_generator('女2', '夏装')
    sheet_generator('女3', '冬装')
    sheet_generator('女3', '春秋')
    sheet_generator('女3', '夏装')
    seperater('男1', '冬装')
    seperater('男1', '春秋')
    seperater('男1', '夏装')
    seperater('男2', '冬装')
    seperater('男2', '春秋')
    seperater('男2', '夏装')
    seperater('男3', '冬装')
    seperater('男3', '春秋')
    seperater('男3', '夏装')
    seperater('女1', '冬装')
    seperater('女1', '春秋')
    seperater('女1', '夏装')
    seperater('女2', '冬装')
    seperater('女2', '春秋')
    seperater('女2', '夏装')
    seperater('女3', '冬装')
    seperater('女3', '春秋')
    seperater('女3', '夏装')
    word('男1', '冬装')
    word('男1', '春秋')
    word('男1', '夏装')
    word('男2', '冬装')
    word('男2', '春秋')
    word('男2', '夏装')
    word('男3', '冬装')
    word('男3', '春秋')
    word('男3', '夏装')
    word('女1', '冬装')
    word('女1', '春秋')
    word('女1', '夏装')
    word('女2', '冬装')
    word('女2', '春秋')
    word('女2', '夏装')
    word('女3', '冬装')
    word('女3', '春秋')
    word('女3', '夏装')





