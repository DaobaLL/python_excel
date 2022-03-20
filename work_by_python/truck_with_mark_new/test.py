from cmath import nan
from operator import index
from openpyxl import load_workbook, Workbook
import numpy as np
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

def copy_imformation_by_head_cell():
    wb = load_workbook('总箱单表-12点00修改.xlsx')
    ws = wb['总箱单']

    f_colum = ws['F']

    f_cells = []
    for x in range(len(f_colum)):
        if f_colum[x].value == None or f_colum[x].value == '':
            f_colum[x].value = f_colum[x-1].value
        f_cells.append(f_colum[x].value)

    myvar = pd.Series(f_cells)
    print(myvar.head(10))

    # 写入读取的f_cells数据
    n=1
    for cell in f_cells:
        ws['F' + str(n)].value = cell
        n += 1
        
    wb.save('总箱单表-12点00修改.xlsx')
    wb.close

# 读取总的箱单数据，并且按照提单号和卡车号排序并出CD申请
def read_infor_and_sort_xlsx():
    infor_by_pd = pd.read_excel('总箱单表-12点00修改.xlsx','总箱单')

    # 按照车头号的值进行排序
    # reset.index() 对排序的df对象重新建立索引

    sorted_df = infor_by_pd.sort_values(by='提单号')
    sorted_df = infor_by_pd.sort_values(by='车头号').reset_index()
    # print(sorted_df.head(50))
    # print(sorted_df.tail(50))
    # print(len(sorted_df))
    # print(type(sorted_df.loc[2000,'车头号']))

    # 用来储存一个CD单证所需卡车信息的index
    TheIndexOf_OnTrucksGoodsInformation = []
    
    # 单证所需卡车index的列表
    list_of_all_CD_that_needed_index =[]

    # 遍历整个表格中的数据，输出{车头号+提单尾号：[信息]} 
    for n in range(0,2000):
        
        # 如果车头号为空，证明货物还没有装车，所以退出循环
        if pd.isna(sorted_df.loc[n,'车头号']) == True:
            break
        
        # 如果上一行车头号、提单号和下一行均相同，则证明这两列的货物装在同一辆车上
        elif sorted_df.loc[n,'车头号'] == sorted_df.loc[n+1, '车头号']:
            if sorted_df.loc[n,'提单号'] == sorted_df.loc[n+1,'提单号']:
                # 车牌号，提单号都相同，继续读取
                TheIndexOf_OnTrucksGoodsInformation.append(n)
            else:
                # 提单号不同，一张CD完成
                TheIndexOf_OnTrucksGoodsInformation.append(n)
                # 出单，调用出单函数
                TheIndexOf_OnTrucksGoodsInformation = []
                list_of_all_CD_that_needed_index.append(TheIndexOf_OnTrucksGoodsInformation)
        else:
            # 车牌号不同，一张CD完成
            TheIndexOf_OnTrucksGoodsInformation.append(n)
            # 出单，调用出单函数
            TheIndexOf_OnTrucksGoodsInformation = []
            list_of_all_CD_that_needed_index.append(TheIndexOf_OnTrucksGoodsInformation)
    
    print(list_of_all_CD_that_needed_index)        


read_infor_and_sort_xlsx()