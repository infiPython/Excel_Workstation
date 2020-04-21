# -*- coding: utf-8 -*-
"""
Created on Thu Apr 16 13:01:08 2020

@author: Haoting Sui

Excel_Merge
Version: 1.0
Desciption: 对Excel文件进行汇总整理
"""

import xlrd, xlwt, os, sys

# 选出所有工作簿共有的工作表
def same_sheets(atts_list):
    sheet_names = xlrd.open_workbook(".\\Attachments\\" + atts_list[0]).sheet_names()
    
    for i in atts_list[1:]:
        book = xlrd.open_workbook(".\\Attachments\\" + i)
        temp_names = book.sheet_names()
        
        removal = 0  # 已移除的元素个数
        for j in range(len(sheet_names)):
            state = False  # 表示是否有相同元素的标志
            
            for k in range(len(temp_names)):                
                if sheet_names[j-removal] == temp_names[k]:
                    state = not state
                    break
                else:
                    pass
            
            if state == False:
                sheet_names.pop(j-removal)
                removal = removal + 1
        
        leftover = removal -len(sheet_names)
        if leftover == 0:
            break
        
    return sheet_names              
            


if __name__ == "__main__":
    atts_list = os.listdir(".\\Attachments")  # 列出需要合并的所有附件 type list
    if atts_list == []:
        print("INFO：【Attachments】文件夹下没有Excel文件")
        input()
        sys.exit(0)
    
    # 定义汇总工作簿及工作表
    new_book = xlwt.Workbook(encoding="utf-8")
    new_sheet = new_book.add_sheet("汇总")
    
    # 列出所有工作表的名称
    book = xlrd.open_workbook(".\\Attachments\\" + atts_list[0])
    sheet_names = book.sheet_names()
    sheets = book.sheets()
    
    # 选择汇总模式
    sheet_names = same_sheets(atts_list)  # 所有工作簿共有的工作表
    if len(sheet_names) > 0:
        print(">>>请选择汇总模式：\n【1】 按索引选择工作表\n【2】 按名称选择工作表")
        mode = input(">>> ")
    else:
        mode = "1"
    
    if mode == "1":  # 按照编号选择工作表    
        # 选择要汇总的工作表
        print(">>>请选择要汇总的工作表的编号：")
        for i in range(len(sheets)):
            print("【" + str(i) + "】")
        sheet_num = int(input(">>> "))
        
        # 定义数据起始行
        print(">>>选择第几行为数据起始行？")
        data_start = int(input(">>> ")) - 1
        
        # 取出所有的信息
        row_to_write = 1
        for i in atts_list:  # 直接循环遍历对象
            book = xlrd.open_workbook(".\\Attachments\\" + i)
            sheet = book.sheets()[sheet_num]
            nrows = sheet.nrows
            ncols = sheet.ncols
            
            # 读取数据内容信息
            for j in range(data_start, nrows):  # 遍历循环行
                row = sheet.row_values(j)
                
                # 写入信息
                for k in range(len(row)):  # 循环遍历列
                    new_sheet.write(row_to_write, k, row[k])
                
                row_to_write = row_to_write + 1
    
    elif mode == "2":  # 按照名称选择工作表        
        # 选择要汇总的工作表
        print(">>>请选择要汇总的工作表的编号：")
        for i in range(len(sheet_names)):
            print("【" + str(i) + "】" + sheet_names[i])
        num = int(input(">>> "))
        sheet_name = sheet_names[num]
        
        # 定义数据起始行
        print(">>>选择第几行为数据起始行？")
        data_start = int(input(">>> ")) - 1
        
        # 取出所有的信息
        row_to_write = 1
        for i in atts_list:  # 直接循环遍历对象
            book = xlrd.open_workbook(".\\Attachments\\" + i)
            sheet = book.sheet_by_name(sheet_name)
            nrows = sheet.nrows
            ncols = sheet.ncols
            
            # 读取数据内容信息
            for j in range(data_start, nrows):  # 遍历循环行
                row = sheet.row_values(j)
                
                # 写入信息
                for k in range(len(row)):  # 循环遍历列
                    new_sheet.write(row_to_write, k, row[k])
                
                row_to_write = row_to_write + 1
                
    # 定义标题行
    print(">>>选择第几行为标题行？")
    head_line = int(input(">>> ")) - 1
    
    # 写入标题行
    head = sheet.row_values(head_line)
    for i in range(len(head)):
        new_sheet.write(0, i, row[i])
    
    # 保存
    new_book.save(".\\汇总结果.xls")