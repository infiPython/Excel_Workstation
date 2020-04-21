# Excel_Seperator
# Version: 3.1
# Desciption: 可以选定工作簿工作表，并按照某列内容进行数据分离（为独立的工作簿或工作表）；同时兼具筛选功能
# Update: 完善功能和代码，可以手动选择标题行和数据开始行，灵活性更强。
# To_do: 能够自动找出标题行及数据行

import xlrd
import xlwt
import os
from shutil import move
from re import sub


# 打印带编号的语句
# 打印带编号的语句
def my_print(num, name="", other=""):
    # num 当前编号 type int
    # name 需要打印的字符串 type str
    print(other + "【" + str(num) + "】 " + name)

def write_record(ncols, record):
    global record_num
    for k in range(ncols):
        sheet_written.write(record_num, k, record[k])
    record_num = record_num + 1  # 更新输入行

def move_rename(filename, filename_new):
    # 需要移动并重命名的文件的名称
    # 不带扩展名的文件名
    suffix = "(1)"
    while True:
        if os.path.exists(".\\" + filename_new + ".xls"):
            filename_new = filename_new + suffix
        else:
            move(".\\Attachments\\" + filename, ".\\" + filename_new + ".xls")
            break


if __name__ == "__main__":

    while True:
        # 选择工作簿
        files = os.listdir(".\\")
        removal = 0
        for i in range(len(files)):
            # 去除非Excel文件
            if ".xls" not in files[i-removal] and ".xlsx" not in files[i-removal]:
                files.remove(files[i-removal])
                removal = removal + 1
        if len(files) == 1:
            print(">>>当前正在处理【" + files[0] + "】")
            filename = files[0]
        else:
            print(">>>请选择要分离的工作簿：")
            for i in range(len(files)):
                my_print(i, files[i])
            sheet_num = int(input(">>> "))
            filename = files[sheet_num]
        
        # 选择工作表
        book = xlrd.open_workbook(".\\" + filename)  # 实例化工作簿
        print(">>>请选择要分离的工作表：")
        sheets = book.sheet_names()
        n = 1
        for i in sheets:
            print("【" + str(n) + "】 " + i)
            n = n + 1
        num = int(input(">>> ")) - 1  # 工作表编号
        sheet = book.sheets()[num]  # 选择工作表
        
        # 定义标题行与数据起始行
        print(">>>请输入标题行所在行位置：")
        head_index = int(input(">>> ")) - 1
        print(">>>请输入数据起始行位置：")
        data_start = int(input(">>> ")) - 1
        
        # 工作表相关信息
        nrows = sheet.nrows  # 计算有效行数
        ncols = sheet.ncols  # 计算有效列数
        sequence = list(range(nrows))[data_start:]  # 数据（包含标题）所在行的索引值
        column_name = []
        print(">>>请选择分离文件的依据列：")
        n = 1
        for i in range(ncols):
            column_name.append(sheet.cell_value(head_index, i))
            print("【" + str(n) + "】" + " " + sheet.cell_value(head_index, i))
            n = n + 1
        num = int(input(">>> ")) - 1  # 分离指标所在列
        
        # 定位模式
        print(">>>请选择定位模式：\n【1】精确定位\n【2】模糊定位")
        loc_mode = input(">>> ")
        if loc_mode == "1":        
            # 寻找类别 【精确定位】
            sorts = []  # 初始化类别
            for i in range(data_start, nrows):
                item = sheet.cell_value(i,num)
                if item not in sorts and item != "":
                    sorts.append(item)
        elif loc_mode == "2":
            # 寻找类别 【模糊定位】
            sorts = []  # 初始化类别
            while True:
                print(">>>搜索：")
                temp = input(">>> ")
                if temp != "":
                    sorts.append(temp)
                else:
                    break
            if sorts == []:
                print("WARNING：类别不能为空")
                input()
                continue
                    
        # 分离模式选择
        # 模式2与模式3性质不同，但操作本质相同
        if len(sorts) > 1:
            print(">>>请选择数据分离模式：\n【1】 分离为多个工作簿\n【2】 分离为多个工作表\n【3】 数据筛选模式")
            mode = input(">>> ")
            print(">>>该工作表数据将会以以下分类进行数据分离：")
            for i in sorts:
                print(i)
            input()
        else:
            mode = "3"
            print(">>>即将进行数据筛选")
            input()
        
        # 挑选数据，生成文件
        if mode == "1":  # 分离为多个工作簿
            try:
                os.mkdir(".\\Attachments")
            except:
                pass
            for i in sorts:  # 分离指标值
                # 创建一个新工作簿
                record_num = 0  # 录入记录数
                book_written = xlwt.Workbook(encoding="utf-8")
                sheet_written = book_written.add_sheet("sheet1")
                record = [sheet.cell_value(head_index, k) for k in range(ncols)]  # 读取标题行
                write_record(ncols, record)  # 写入标题行
                
                for j in range(len(sequence)):  # 被录入记录所在行索引
                    index = sheet.cell_value(sequence[j-record_num], num)
                    
                    # 【精确定位】 数据分离
                    if loc_mode == "1":
                        if i == index:
                            record = [sheet.cell_value(sequence[j-record_num], k) for k in range(ncols)]  # 读取一条记录                      
                            sequence.pop(j-record_num)
                            # 录入一条数据
                            write_record(ncols, record)
                        else:
                            continue
                    # 【模糊定位】 数据分离
                    elif loc_mode == "2":
                        if i in index:
                            record = [sheet.cell_value(sequence[j-record_num], k) for k in range(ncols)]  # 读取一条记录                      
                            sequence.pop(j-record_num)
                            # 录入一条数据
                            write_record(ncols, record)
                        else:
                            continue
                        
                book_written.save(".\\" + "Attachments" + "\\" + i + ".xls")        
        
        elif mode == "2" or mode == "3":  # 分离为多个工作表
            book_written = xlwt.Workbook(encoding="utf-8")
            for i in sorts:  # 分离指标值
                sheet_written = book_written.add_sheet(i)
                record = [sheet.cell_value(head_index, k) for k in range(ncols)]  # 读取标题行
                record_num = 0  # 录入记录数
                write_record(ncols, record)  # 写入标题行
                
                for j in range(len(sequence)):  # 被录入记录所在行索引
                    index = sheet.cell_value(sequence[j-record_num], num)
                    
                    # 【精确定位】 数据分离
                    if loc_mode == "1":
                        if i == index:
                            record = [sheet.cell_value(sequence[j-record_num], k) for k in range(ncols)]  # 读取一条记录
                            sequence.pop(j-record_num)
                            # 录入一条数据
                            write_record(ncols, record)
                        else:
                            pass
                    # 【模糊定位】 数据分离
                    elif loc_mode == "2":
                        if i in index:
                            record = [sheet.cell_value(sequence[j-record_num], k) for k in range(ncols)]  # 读取一条记录
                            sequence.pop(j-record_num)
                            # 录入一条数据
                            write_record(ncols, record)
                        else:
                            pass
            
            try:
                os.mkdir(".\\Attachments")
            except:
                pass
            filename = sub(".xlsx", ".xls", filename)
            book_written.save(".\\" + "Attachments" + "\\" + filename)
        
        
        else:  # 输入非法指令
            print(">>>输入非法指令，正在重启...\n\n")
            continue
        
        # 对结果进行汇报处理
        if mode == "1":
            print(">>>数据分离已完成，请查看【Attachments】文件夹")
            input(">>>回车进入下一个工作流程\n\n")
        elif mode == "2":
            print(">>>数据分离已完成")
            move_rename(filename, os.path.splitext(filename)[0])
            input(">>>回车进入下一个工作流程\n\n")
        else:
            filename_new = "筛选数据"
            move_rename(filename, filename_new)
            print(">>>数据筛选已完成")
            input(">>>回车进入下一个工作流程\n\n")