# -*- coding: utf-8 -*-
"""
Created on Thu Apr 16 18:11:22 2020

@author: Haoting Sui

# Excel_Preview
# Version: 1.0
# Description: 快速预览一组工作簿的工作表
"""
import xlrd, os, sys

def read_config():
    book = xlrd.open_workbook(".\\Configuration\\Preview_Configuration.xlsx")
    sheet = book.sheets()[0]
    display_len = int(sheet.cell_value(1,0))
    rows_max = int(sheet.cell_value(1,1))
    
    return display_len, rows_max
    
# 寻找相同的工作表
def same_sheets(atts_list):
    # Excel文件类别 type list
    com_sheet_names = xlrd.open_workbook(".\\Attachments\\" + atts_list[0]).sheet_names()
    
    for i in atts_list[1:]:
        book = xlrd.open_workbook(".\\Attachments\\" + i)
        temp_names = book.sheet_names()
        
        removal = 0  # 已移除的元素个数
        for j in range(len(com_sheet_names)):
            state = False  # 表示是否有相同元素的标志
            
            for k in range(len(temp_names)):                
                if com_sheet_names[j-removal] == temp_names[k]:
                    state = not state
                    break
                else:
                    pass
            
            if state == False:
                com_sheet_names.pop(j-removal)
                removal = removal + 1
        
        leftover = removal -len(com_sheet_names)
        if leftover == 0:
            break
        
    return com_sheet_names  # 各个工作簿所共有的工作表的名称 type list

# 打印一条记录
def print_record(record, display_len, ncols, row_index):
    # record 需要打印的一条记录（长度已经过修剪，小于display_len） type list
    # display_len 每一个单元格显示的最大长度 引用自Configuration type int
    # ncols 工作表的列数
    # row_index 当前写入行的行索引值
    
    sys.stdout.write("【{}】|".format(row_index))
    for i in range(ncols):
        print_format = " {" + "0:-<{}".format(display_len) + "} |"
        sys.stdout.write(print_format.format(record[i]))
    print("\n")

# 修改一条记录
def modify_record(record,display_len):
    # record 一条记录的信息 type list
    # display_len 每一个单元格显示的最大字符数 type int
    
    for i in range(len(record)):
        if len(str(record[i])) > display_len:
            record[i] = str(record[i])[0:display_len]
    return record

# 预览工作表
def preview(display_len, rows_max, book_name):
    # rows_max 预览显示的最大行数 【type int】 引用自Configuration
    # display_len 单元格显示的最大长度 【type int】 引用自Configuration
    global sheet
    nrows = sheet.nrows
    ncols = sheet.ncols
    
    # 打印出一个工作表的【row_index】条记录的信息
    print("【{}】".format(book_name))
    rows_max = rows_max if rows_max < nrows else nrows
    for i in range(rows_max):        
        record = sheet.row_values(i)
        record = modify_record(record, display_len)  # 对单元格内容进行修剪
        print_record(record, display_len, ncols, i+1)
    print()


if __name__ == "__main__":
    atts_folder = ".\\Attachments"  # 工作簿所在文件夹 type str
    atts_list = os.listdir(".\\Attachments")  # 需要预览的文件列表 type list
    display_len, rows_max = read_config()  # 读取程序配置，返回【display_len】单元格显示长度、【rows_max】最大显示行数
    display_len = int(display_len)
    
    if atts_list == []:
        print("INFO：当前文件夹下无Excel文件...")
        input()
        sys.exit(0)
    
    # 模式选择
    com_sheet_names = same_sheets(atts_list)  # 找出共有的工作表的名称 type list
    while True:
        if com_sheet_names != []:
            print(">>>请选择预览模式：\n【1】 预览表名相同的工作表\n【2】 按编号预览工作表")
            mode = input(">>> ")
            
            if mode == "2":
                sheet_num = 0
                book = xlrd.open_workbook(os.path.join(atts_folder, atts_list[0]))
                sheet_num = len(book.sheets()) 
                for i in range(len(atts_list)):
                    book_name = atts_list[i]
                    book = xlrd.open_workbook(os.path.join(atts_folder, book_name))
                    length = len(book.sheets())
                    if sheet_num > length:
                        sheet_num = length
                    
        else:
            mode = "2"
            
            book = xlrd.open_workbook(os.path.join(atts_folder, atts_list[0]))
            sheet_num = len(book.sheets())        
            for i in range(len(atts_list)):
                book_name = atts_list[i]
                book = xlrd.open_workbook(os.path.join(atts_folder, book_name))
                temp = len(book.sheets())
                if sheet_num > length:
                    sheet_num = length
        
        # 各工作簿有名称相同的工作表
        if mode == "1":      
            print(">>>请选择需要预览的工作表")
            for i in range(len(com_sheet_names)):
                print("【{}】 {}".format(str(i), com_sheet_names[i]))
            sheet_no = int(input(">>> ")) - 1
            
            for i in range(len(atts_list)):
                book_name = atts_list[i]
                book = xlrd.open_workbook(os.path.join(atts_folder, book_name))
                sheet = book.sheet_by_name(com_sheet_names[sheet_no])
                
                preview(display_len, rows_max, book_name)
                
        elif mode == "2":
            print(">>>请选择需要预览的工作表")
            for i in range(sheet_num):
                print("【{}】".format(str(i)))
            sheet_no = int(input(">>> "))
            
            for i in range(len(atts_list)):
                book_name = atts_list[i]
                book = xlrd.open_workbook(os.path.join(atts_folder, book_name))
                sheet = book.sheets()[sheet_no]
                nrows = sheet.nrows
                
                preview(display_len, rows_max, book_name)
        
        print(">>>是否继续在当前工作簿下继续？\n【{0:^6}】继续\n【{1:^6}】结束".format("1", "Return"))
        choice = input(">>> ")
        if choice == "1":
            pass
        elif choice == "":
            break