import xlwings as xw
import pickle
import os
import math
import copy
wb=xw.Book(r'高中英语单词检索词汇总表(人教版)(必修1至选修8).xlsm')
sht=wb.sheets['Sheet1']
sht2=wb.sheets['Sheet2']
sht3=wb.sheets['Sheet3']
data1=sht.range((1,3),(3000,3)).value#需要背诵的全部单词
data2=sht2.range((1,1),(3000,1)).value#需要背诵的全部单词
data3=sht3.range((1,1),(3000,1)).value#需要背诵的全部单词
for a in range(3000):
    print('a:%d' %a)
    if data3[a]:
        if list(data3[a])[0]=='△':
            for b in range(3000):
                if data1[b]:
                    if data1[b].lower()==''.join(list(data3[a])[1:]).lower():
                        sht.cells(b+1,9).value=1
wb.save()
    