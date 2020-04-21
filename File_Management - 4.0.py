# -*- coding: utf-8 -*-
"""
Created on Sat Apr 18 00:30:05 2020

@author: Haoting Sui

# Files_Management
# Version: 4.0
# Description: 对指定的文件夹及其子文件夹【和子文件夹的子文件夹...】（包括压缩文件）中的文件进行【复制】、【移动】、【查找】、【查看】
# Update: 1、搜索范围扩大到非系统磁盘 2、解决了查找逻辑中的一个bug
# To_do: 1、解压缩文件名称乱码的问题,添加自动解压缩开关
"""

import os, shutil, re


# 挑选出文件夹
def select_folder(local_path=".\\"):
    # local_path 【files_list】中的文件所在的相对目录
    folders_list = []
    
    previous_dir = os.getcwd()  # 记录原来的路径
    os.chdir(local_path)  # 更改当前路径
    for i in os.listdir(".\\"):
        if os.path.isdir(i):
            folders_list.append(i)
    
    os.chdir(previous_dir)
    
    return folders_list

# 挑选出文件
def select_files(num1, num2):
    # num1 符合条件的文件数目 type int
    # num2 符合条件的文件夹数目 type int
    state1 = 0
    state2 = 0
    num = num1 + num2
    serial_num = []
    
    if num > 0:
        print(">>>请选择需要找出的文件及文件夹：")
        print("【{0:^8}】 跳过\n【{1:^8}】 全选\n【{2:^8}】 挑选".format("NULL", "0", "$No.$"))
        while True:
            choice = input(">>> ")
            if choice == "0":
                serial_num = list(range(1, num+1))
                state1 = state1 or 1
                if num2 != 0:
                    state2 = state2 or 1
                return serial_num, state1, state2
            elif choice == "":
                return serial_num, state1, state2
            else:
                if int(choice) <= num1:
                    state1 = state1 or 1
                else:
                    state2 = state2 or 1
                serial_num.append(choice)
    else:
        return serial_num, state1, state2

# 创建文件夹
def mk_dir(target_path):
    # target_path 目标路径
    suffix = "(1)"
    
    while True:
        if os.path.exists(target_path):
            target_path = target_path + suffix
        else:
            os.mkdir(target_path)
            break
    
    return target_path

# 在文件列表中查找，将符合条件的赋予编号存入新的列表
def f_match(all_f, features, no_start):
    # all_f 所有的文件或文件夹 type list
    # features 文件或文件夹需要满足的特征 type list
    # no_start 编号起始值 type int
    target_f = {}  # 文件或文件夹的匹配值 type dict
    match = []  # 匹配的关键字
    unmatch = []  # 不匹配的关键字
    
    all_f_temp = all_f
    temp = []
    state = 0  # 【feature】匹配状态，初始化为0，未匹配
    num = 0
    for f in all_f_temp:  # 第一个参数在文件或文件夹名中匹配，【因此第一个参数极其重要，有利于提高准确度和缩小范围】
        if features[0] in f[1].lower():
            temp.append([f[0], f[1]])
            state = state or 1
            num = num + 1
    if state == 1:
        match.append(features[0])
    else:
        unmatch.append(features[0])
    for feature in features[1:]:  # 第二个至若干个参数在全路径中匹配
        all_f_temp = temp
        temp = []
        state = 0
        num = 0
        for f in all_f_temp:
            if feature in (f[0]+f[1]).lower():
                temp.append([f[0], f[1]])
                state = state or 1
                num = num + 1
        if state == 1:  # 特征匹配
            match.append(feature)
        else:  # 特征不匹配
            unmatch.append(feature)
    
    for i in range(num):
        target_f[str(i+1+no_start)] = temp[i]
    
    return target_f, match, unmatch, num
    
# 分离非压缩文件和压缩文件
def classify(files_list):
    # files_list 当前文件夹下的所有文件的文件名 type list

    compressed_files_dict = {}  # 压缩文件
    transferred_num = 0
    
    for i in range(len(files_list)):
        match = re.search("(?:(?:.zip)|(?:.rar)|(?:.tar)|(?:.gz))$", files_list[i-transferred_num])
        if match:
            compressed_files_dict[files_list.pop(i-transferred_num)] = match.group()  # 将压缩文件从【files_list】中，并将文件名及压缩格式添加到【compressed_files】中去
            transferred_num = transferred_num + 1
    return files_list, compressed_files_dict

# 【递归】处理非压缩文件及文件夹
def find_out(file_or_folder, local_path):
    # file_or_folder 文件名（带扩展名）或文件夹名称 type str
    # local_path 【file_or_folder】所在的文件目录，默认为程序所在的根目录 type str
    global local_sym, all_files, all_folders
    
    temp = os.path.join(local_path, file_or_folder)  # 【file_or_folder】的路径
    if os.path.isdir(temp):  # 如果该文件是一个文件夹
        local_path = temp
        files_list = os.listdir(local_path)  # 列出文件夹中所有的文件
        folders = select_folder(local_path)
        all_folders.extend([[local_path, i] for i in folders])  # 添加当前目录下所有的文件夹
        # 将文件夹写在文件列表最后
        n = 0
        for i in range(len(files_list)):
            for j in range(len(folders)):
                if files_list[i-n] == folders[j]:
                    files_list.pop(i-n)
                    n = n + 1
                    break
        files_list.extend(folders)
        # 递归
        for i in files_list:
            find_out(i, local_path)  # 【递归的入口】            
            
    else:  # 如果不是一个文件夹，就移到汇总文件夹中
            all_files.append([local_path, file_or_folder])
        
        
if __name__ == "__main__":
    local_sym = ":\\"  # 表示【当前文件夹】的符号，同时也是根目录
    target_path = os.path.join(local_sym, "Attachments") # 目标文件夹的绝对路径 type str
    mapping = {"1": "d:\\", "2": "e:\\", "3": "g:\\"}
    print(">>>请选择磁盘\n【1】 D盘\n【2】 E盘\n【3】 G盘")
    disk = mapping[input(">>> ")]
    
    while True:
        folders_list = select_folder(disk)  # 程序所在目录的文件夹 type list    
        all_files = []  # 选定文件夹下的所有文件（不包含文件夹）的所在目录及名称，是一个二维数组列表 type list
        all_folders = []  # 选定文件夹下的所有文件夹所在的目录及名称，是一个二维数组，type list
        
        # 选择文件夹
        f_list_len = len(folders_list)
        if f_list_len == 0:
            print(">>>当前目录无可处理的文件夹")
        elif f_list_len == 1:
            print(">>>当前正在处理【" + folders_list[0] + "】")
            input()
            file_or_folder = folders_list[0]
        else:
            print(">>>请选择文件夹")
            for i in range(f_list_len):
                print("【" + str(i) + "】" + folders_list[i])
            file_no = int(input(">>> "))
            file_or_folder = folders_list[file_no]
        
        # 选择处理模式
        print(">>>请选择处理模式：")
        print("【1】 复制\n【2】 移动\n【3】 查找\n【4】 查看")
        mode = input(">>> ")
        find_out(file_or_folder, disk)
        
        # 执行处理
        all_files_num = len(all_files)
        if mode == "1":  # 复制
            mk_dir(target_path)
            for file in all_files:
                shutil.copyfile(os.path.join(file[0], file[1]), os.path.join(target_path, file[1]))
            print("INFO：共{}个文件已全部复制到【{}】目录下".format(all_files_num, target_path))
            input()
            
        elif mode == "2":  # 移动
            mk_dir(target_path)
            for file in all_files:
                shutil.move(os.path.join(file[0], file[1]), os.path.join(target_path, file[1]))
            print("INFO：共{}个文件已全部移动到【{}】目录下".format(all_files_num, target_path))
            input()
            
        elif mode == "3":  # 查找
            while True:
                print(">>>请输入文件特征（不区分大小写）：")
                features = []  # 文件特征 type list
                # 收集文件特征
                while True:
                    feature = input(">>> ")
                    if feature != "":
                        features.append(feature)
                    else:
                        break
                features_match = [[[], []], [[], []]]  # 特征匹配结果
                
                # 匹配文件
                target_files, features_match[0][0], features_match[0][1], num1 = f_match(all_files, features, 0)
                # 匹配文件夹
                target_folders, features_match[1][0], features_match[1][1], num2 = f_match(all_folders, features, num1)
                
                # 列出符合条件的文件及文件夹
                print(">>>有以下文件匹配：\n")
                print("INFO：【文件】 匹配：{} | 未匹配：{}\n".format("、".join(features_match[0][0]), "、".join(features_match[0][1])))
                for (i, file) in target_files.items():
                    print(re.sub(r"\\", r" \\ ", "【{}】 {}\【{}】\n".format(i, file[0], file[1])))
                print("\nINFO：【文件夹】 匹配：{} | 未匹配：{}\n".format("、".join(features_match[1][0]), "、".join(features_match[1][1])))
                for (i, folder) in target_folders.items():
                    print(re.sub(r"\\", r" \\ ", "【{}】 {}\【{}】\n".format(i, folder[0], folder[1])))
                
                # 选择文件及文件夹
                serial_num, state1, state2 = select_files(num1, num2)  # 需要挑出的文件及文件夹的编号 type list
                
                # 复制匹配的项目至目标文件夹
                if state1 == 1:  # 有文件需要被复制至目标路径中
                    target_path_files = mk_dir(".\\【文件】"+"、".join(features_match[0][0]))
                if state2 == 1:  # 有文件夹需要被赋值至目标路径中
                    target_path_folders = mk_dir(".\\【文件夹】"+"、".join(features_match[1][0]))
                for i in serial_num:
                    if int(i) <= num1:
                        # 复制文件
                        shutil.copyfile(os.path.join(target_files[i][0], target_files[i][1]), os.path.join(target_path_files, target_files[i][1]))
                    else:
                        # 复制文件夹
                        shutil.copytree(os.path.join(target_folders[i][0], target_folders[i][1]), os.path.join(target_path_folders, target_folders[i][1]))
                
                print(">>>是否在【{}】文件夹中继续查找？\n【NULL】 是\n【0】 否".format(file_or_folder))
                choice = input(">>> ")
                if choice == "":
                    pass
                elif choice == "0":
                    break
        
        elif mode == "4":  # 查看
            num = 0
            for file in all_files:
                print(re.sub(r"\\", r" \\ ", "{}\【{}】\n".format(file[0], file[1])))
                num = num + 1
            print("INFO：共计{}个文件".format(num))
            input()