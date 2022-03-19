import imp
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


infor_by_pd = pd.read_excel('总箱单表-12点00修改.xlsx','总箱单')
# print(infor_by_pd.head(10))


# for x in infor_by_pd:
#     print(x)


# print(len(infor_by_pd))

# for x in range(0,len(infor_by_pd)):
#     print(infor_by_pd.loc[x,'车头号'])


for n in range(6,16):
    print(infor_by_pd.loc[n,'箱件数  PKGs Qt\'y'])
    
    print(infor_by_pd.loc[n,'规格型号  Specification'])

# 按照车头号的值进行排序
sorted_df = infor_by_pd.sort_values(by='车头号')
print(sorted_df.describe)
print(type(infor_by_pd))
# 
