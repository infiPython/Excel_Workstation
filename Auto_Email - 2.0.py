# -*- coding: utf-8 -*-
"""
Created on Sun Apr 19 19:34:09 2020

@author: Haoting Sui

# Auto_Email
# Version: 2.1
# Desciption: 批量点对点发送邮件
# Update: 1、进行参数修改，与Excel_Seperator Excel_Merge进行协同配合 2、修复了发送中文名称附件时文件名更改及后缀变成“bin”的错误
# To do: 对程序显示逻辑进行修改
"""

import smtplib, xlrd, xlwt, os, shutil, sys
from email.mime.text import MIMEText  # 文本
from email.mime.multipart import MIMEMultipart  # 附件
from email.header import Header
from time import sleep


def read_config():
    book = xlrd.open_workbook(".\\Configuration\\Email_Configuration.xlsx")
    sheet = book.sheets()[0]

    variables = sheet.col_values(0)
    values = sheet.col_values(1)
    config = dict(map(lambda x,y:[x,y],variables, values))

    return config

def read_task():
    book = xlrd.open_workbook(".\\Configuration\\邮件发送列表.xlsx")
    sheet = book.sheets()[0]

    name = sheet.col_values(0)[1:]
    email_addr = sheet.col_values(1)[1:]
    atts = sheet.col_values(2)[1:]

    return name, email_addr, atts

def add_content(content,from_who, to_who, subject):
    global message
    
    # 邮件设置
    message["From"] = Header(from_who, "utf-8")
    message["To"] = Header(to_who, "utf-8")
    message["Subject"] = Header(subject, "utf_8") 

def add_attachment(content, att_addr, from_who, to_who, subject):
    # content 邮件正文
    # att_addr 附件全称
    # from_who 发件人备注
    # to_who 收件人备注
    global message
    
    # 邮件设置
    message["From"] = Header(from_who, "utf-8")
    message["To"] = Header(to_who, "utf-8")
    message["Subject"] = Header(subject, "utf_8")

    # 编辑邮件正文
    message.attach(MIMEText(content, "plain", "utf-8"))
    # 添加附件
    att1 = MIMEText(open(".\\Attachments\\" + att_addr, "rb").read(), "base64", "utf-8")
    att1["Content-Type"] = "application/octet-stream"
    
    # att1["Content-Disposition"] = 'attachment; filename="' + att_addr + '"'  不支持中文
    att1.add_header('Content-Disposition', 'attachment', filename=att_addr)
    message.attach(att1)

def att_mapping(receivers_name, receivers_addr, atts):
    # receivers_name 从邮件发送列表中读取的收件人备注 type list
    # receivers_addr 从邮件发送列表中读取的收件人邮箱地址 type list
    # atts 从邮件发送列表中读取的附件名称 type list
    
    atts_list = os.listdir(".\\Attachments")  # 实际需要发送的附件的全称
    atts_dict = {}  # 任务信息组合
    atts_index = []  # 需要发送的附件名称
    
    # 将任务信息进行组合
    for i in range(len(atts)):
        if atts[i] != "":
            atts_index.append(atts[i])
            atts_dict[atts[i]] = [receivers_name[i], receivers_addr[i]]
    
    # 在任务信息中加入附件全名
    for i in range(len(atts_list)):
        for j in range(len(atts)):
            if atts[j] in atts_list[i]:
                atts_dict[atts[j]].append(atts_list[i])
                atts.pop(j)
                break
    
    return atts_dict, atts_index  # 邮件任务的组合、需要发送的附件名称


if __name__ == "__main__":    

    # 读取配置信息
    config = read_config()
    mail_host = config["mail_host"]  # 邮箱主机
    port = config["port"]  # 服务器端口号
    mail_user = config["mail_user"]  # 邮箱用户名
    mail_pass = config["mail_pass"]  # 邮箱授权码
    sender = config["sender"]  # 发件人邮箱地址
    
    # 登录邮箱
    smtpObj = smtplib.SMTP()
    try:
        smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
    except:
        print("INFO：连接失败！请检查网络连接设置或查看SMTP端口设置是否正确...")
        input()
        sys.exit(0)
    try:
        smtpObj.login(sender, mail_pass)
    except:
        print("INFO：连接失败！请检查发件人邮箱地址输入是否有误，或授权码是否有效...")

    # 读取邮件发送任务
    receivers_name, receivers_addr, atts = read_task()  # 收件人备注、收件人地址、附件名称
    atts_dict, atts_index = att_mapping(receivers_name, receivers_addr, atts)  # 邮件任务的组合、需要发送的附件名称
    
    while True:
        # 模式选择
        print(">>>请选择发送内容：")
        print("【1】 文本\n【2】 附件")
        mode = input(">>> ")
    
        # 定义邮件发送结果变量
        success = 0
    
        # 编辑文本邮件
        subject = input(">>>请输入邮件主题：")
        if mode == "1":
            content = input(">>>请输入要发送的正文内容：")
            try:
                for i in range(len(receivers_addr)):  # 编辑每个收件人的邮件内容
                    message = MIMEText(content, "plain", "utf-8")  # 创建一个邮件实例
                    add_content(content, mail_user, receivers_name[i], subject)
                    # 发送邮件
                    smtpObj.sendmail(sender, receivers_addr[i], message.as_string())
                    print("INFO：邮件已发送至【" + receivers_name[i] + "】。")
                    success = success + 1
                    sleep(0.5)
            except:
                print("ERR：至【" + receivers_name[i] + "】的邮件发送失败。")
    
        # 编辑附件邮件
        elif mode == "2":
            content = input(">>>请输入要发送的正文内容：")
            for i in atts_index:
                message = MIMEMultipart()  # 创建一个带附件的邮件实例
                try:
                    add_attachment(content, atts_dict[i][2], mail_user, atts_dict[i][0], subject)
                except:
                    print("ERR：未找到附件【" + i + "】")
                    continue
                # 发送邮件
                smtpObj.sendmail(sender, atts_dict[i][1], message.as_string())
                print("INFO：附件【" + atts_dict[i][2] + "】" + "已发送至【" + atts_dict[i][0] + "】")
                # 将成功发送的邮件转移至Success文件夹
                shutil.move(".\\Attachments\\" + atts_dict[i][2], ".\\Success\\" + atts_dict[i][2])
                success = success + 1
                sleep(0.5)
    
        
        # 邮件发送结果报告
        if len(atts_index)-success == 0:
            print(">>>邮件已全部发送成功，回车以继续...")
            input()
        else:
            print(">>>共计：" + str(success) + "人发送成功；" + str(len(atts_index)-success) + "人未发送，回车以继续...")
            input()