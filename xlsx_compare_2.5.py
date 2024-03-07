#!/usr/bin/env python
# -*- coding:utf-8 -*-
# filename: xlsx_compare
# author: zhong

# date: 2023/11/10
# 增加格式输出的整洁调整
# date: 2023/10/20
# 增加sheet数量不一致时的检查
# date: 2023/05/11
# 增加数量不一致时的提示功能
# 增加多个sheet对比功能
# date: 2023/04/25
# 增加A、B内多个文件夹对比功能, 保留直接对比excel文件功能
# date: 2023/04/10
# 增加进度条功能

from pandas import read_excel
import os
from pandas.testing import assert_frame_equal
from tqdm import trange

# 核对信息设置
def my_log(info):
    try:
        with open('核对_info.txt', 'w+', encoding="UTF-8") as f:
            f.writelines(info)
    except Exception as e:
        print('写入错误日志时发生以下错误：\n%s'%e)

# 读取文件
def compare_xlsx():
    try:
        # 获取2个文件夹下的文件, 直接放入文件或文件夹都可以      
        curr_path = os.getcwd()
        A_path = os.path.join(curr_path, 'A')
        B_path = os.path.join(curr_path, 'B')
        
        A_files_name = []        
        A_files = []        
        for paths, dirs, files in os.walk(A_path):
            for f in files:
                if f.endswith('xlsx'):
                    A_files_name.append(f)
                    A_files.append(os.path.join(paths,f))
        B_files_name = []        
        B_files = []        
        for paths, dirs, files in os.walk(B_path):
            for f in files:
                if f.endswith('xlsx'):
                    B_files_name.append(f)
                    B_files.append(os.path.join(paths,f))

        err_lst_0 = ['A、B中，文件数量不一致, 请检查\n\n']
        if len(A_files_name) != len(B_files_name):
            err_lst_0_1 = [x for x in B_files_name if x not in A_files_name]
            err_lst_0_2 = [x for x in A_files_name if x not in B_files_name]
            err_lst_0 = str(err_lst_0[0]) + 'A没有:\n' + '\n'.join(err_lst_0_1) + '\n' + 'B没有:\n' + '\n'.join(err_lst_0_2)            
            my_log(err_lst_0)
        else:
            err_lst_1 = ['A、B中，文件名不一致, 请检查\n\n']
            for i in range(len(A_files_name)):
                if A_files_name[i] != B_files_name[i]:
                    err_lst_1.append('A: ' + A_files_name[i] + '\n' + 'B: ' +  B_files_name[i] + '\n')
            if len(err_lst_1) > 1:
                my_log(err_lst_1)   
        
        # 如果文件数量，名称都没问题，则挨个儿对比文件内容
        if len(err_lst_0) == 1 and len(err_lst_1) == 1:
            err_lst_2 = ['A、B中，文件内容不一致, 请检查'+'\n'*3]
            err_lst_3 = ['文件内sheet数量不一致, 请检查'+'\n'*3]
            for i in trange(len(A_files), desc = "excel进度", unit = '个表', colour='green'):
                df_A_sheet_name = read_excel(A_files[i], na_values='...', sheet_name=None, nrows=1).keys()
                df_B_sheet_name = read_excel(B_files[i], na_values='...', sheet_name=None, nrows=1).keys()
                df_A_sheet_n = len(df_A_sheet_name)
                df_B_sheet_n = len(df_B_sheet_name)
                if df_A_sheet_n != df_B_sheet_n:
                    if df_A_sheet_n > df_B_sheet_n:
                        err_lst_3.append(A_files[i]+'\n'+B_files[i]+'\n'+'差异：'+str(df_A_sheet_name - df_B_sheet_name)+'\n')
                    else:
                        err_lst_3.append(A_files[i]+'\n'+B_files[i]+'\n'+'差异：'+str(df_B_sheet_name - df_A_sheet_name)+'\n')                        
                else:
                    df = read_excel(A_files[i], na_values='...', sheet_name=None, nrows=1) # read only 1 row, speed up for get sheet name
                    for s in df.keys():
                        df_1 = read_excel(A_files[i], na_values='...', sheet_name=s)
                        df_2 = read_excel(B_files[i], na_values='...', sheet_name=s)
                        try:
                            assert_frame_equal(df_1, df_2)
                        except AssertionError as error:
                            err_lst_2.append(A_files[i]+'\n' + B_files[i]+'\n' + f'sheet_name: {s}\n' + str(error)+'\n'+'-'*42*2+'\n'*2)
            
            if len(err_lst_3) > 1:
                my_log(err_lst_3)
            elif len(err_lst_2) > 1:
                my_log(err_lst_2)
            else:
                my_log('一切正常')                

    except Exception as e:
        my_log(str(e))

if __name__ == '__main__':
    with open('核对_info.txt', 'w', encoding="UTF-8") as f:
        f.writelines('\n')
    compare_xlsx()

os.system('start "" "核对_info.txt"')
