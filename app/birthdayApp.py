#import smtplib   ##
import pandas as pd    #时间处理
import datetime        #用于获取时间
from datetime import date,timedelta
import xlrd            #用于处理excel表
from chinese_calendar import is_workday     #用于计算节假日

from tkinter import *
import tkinter as tk

import random
from PIL import ImageTk, Image
import numpy as np

# 读取Excel表格
#df = pd.read_excel("employee.xlsx")
#r'D:\Users\liujintong\PycharmProjects\birthday_celebrition\venv\employee.xlsx'
#读取excel表
file = xlrd.open_workbook(r'related_files/employee.xlsx')
# file = xlrd.open_workbook(r'related_files/employee_formal.xlsx')
sheet = file.sheet_by_index(0)

sheet_rows = sheet.nrows   #excel表总行数
sheet_cols = sheet.ncols   #excel表总列数

#读取当天时间
today = datetime.date.today()
today_str = today.strftime("%m-%d")  #转换成字符串格式


#tomorrow_str = today + timedelta(days=1)
#yestoday = today - timedelta(days=1)
#yestoday_str = yestoday.strftime("%Y-%m-%d")
#today_year = today.strftime("%Y")
name_all = ''

#计算当天是否有人生日
birthday_flag = 0

#生日人数
birthday_emploee_num = 0
def today_birthday_people():
    global birthday_flag,birthday_emploee_num,name_all

    name_all = ''
    birthday_emploee_num = 0

    for i in range(1, sheet_rows):  # 1-sheet_rows，但不包括sheet_rows

        birthday_str = xlrd.xldate.xldate_as_datetime(sheet.cell(i, 1).value, 0).strftime("%m-%d")
        if today_str == birthday_str:
            name = sheet.cell(i, 0).value
            birthday_emploee_num = birthday_emploee_num + 1
            if name_all == '':
                name_all = name
            elif name_all != '':
                name_all = name_all + ' ' + ' ' + name
            print(name, "'s birthday is today! [%s]" % today_str)
            birthday_flag = 1
            continue
        elif today_str != birthday_str:
            continue

    if birthday_flag == 0:
        print("nobody's birthday is today [%s]" % today_str)
    elif birthday_flag == 1:
        birthday_flag == 0

today_birthday_people()

#birthday_str = xlrd.xldate.xldate_as_datetime(sheet.cell(1,1).value, 0).strftime("%m-%d")

#print(birthday_str)
#name = sheet.cell(1,0)
#print(name)

#######################################################

def vocation_birthday_people():
    global name_all,Name_text,birthday_emploee_num
    # 提取当天的年，月，日，后续用于判断是否是假期

    # 读取当天时间
    today = datetime.date.today()
    today_str = today.strftime("%m-%d")  # 转换成字符串格式

    today_year = today.strftime("%Y")
    today_month = today.strftime("%m")
    today_day = today.strftime("%d")
    today_year_int = int(today_year)
    today_month_int = int(today_month)
    today_day_int = int(today_day)
    today_date = datetime.datetime(today_year_int, today_month_int, today_day_int)

    # tomorrow = today + timedelta(days=1)

    if is_workday(today_date):
        print("今天是%s" % today_str)
        holiday_number = 1
        flag = True
        while flag:
            tomorrow = today + timedelta(days=holiday_number)
            tomorrow_str = tomorrow.strftime("%m-%d")
            tomorrow_year = tomorrow.strftime("%Y")
            tomorrow_month = tomorrow.strftime("%m")
            tomorrow_day = tomorrow.strftime("%d")
            tomorrow_year_int = int(tomorrow_year)
            tomorrow_month_int = int(tomorrow_month)
            tomorrow_day_int = int(tomorrow_day)
            tomorrow_date = datetime.datetime(tomorrow_year_int, tomorrow_month_int, tomorrow_day_int)
            if is_workday(tomorrow_date):
                # print("明天是工作日")
                print("%s是工作日" % tomorrow_str)
                flag = False
                #
                # holiday_finnal_day = tomorrow + timedelta(1)
                # holiday_finnal_day_str = tomorrow.strftime("%m-%d")
                #
            else:
                print("%s是休息日" % tomorrow_str)
                #
                for i in range(1, sheet_rows):  # 1-sheet_rows，但不包括sheet_rows
                    birthday_str = xlrd.xldate.xldate_as_datetime(sheet.cell(i, 1).value, 0).strftime("%m-%d")
                    if tomorrow_str == birthday_str:
                        name = sheet.cell(i, 0).value
                        birthday_emploee_num = birthday_emploee_num + 1
                        if name_all == '':
                            name_all = name
                        elif name_all != '':
                            name_all = name_all + ' ' + ' ' + name
                        print(name, "'s birthday is %s!" % tomorrow_str)
                        continue
                    elif tomorrow_str != birthday_str:
                        continue
                #
                holiday_number = holiday_number + 1

    else:
        print("今天是休息日[%s] %today_str")

    # date = datetime.datetime(2023, 7, 22)
    # if is_workday(date):
    #  print("是工作日")
    # else:
    #  print("是休息日")

    # Move_text = "[18:00] 测试测试测试测试测试测试测试测试~"  # 弹幕文本内容
    if name_all == '':
        Name_text = "今天没有人过生日哦~"
    else:
        # Name_text = name_all + ' , Happy Birthday!'
        Name_text = name_all


vocation_birthday_people()

def load_pictures():
    global image_default,image
    # 加载图片
    # image = PhotoImage(file="image.png")
    emploee_picture_array = birthday_name_array = name_all.split('  ')

    # length = len(birthday_name_array)

    image = []
    for i in range(birthday_emploee_num):
        image.append(0)

    if birthday_emploee_num == 0:
        image_png = 'related_files/photos/default_image.jpg'
        image_default = Image.open(image_png)
        resized_img = image_default.resize((700, 275))
        image_default = ImageTk.PhotoImage(resized_img)
        # image_default = PhotoImage(file=image_png)
        # 将缩放后的图片转换为Tkinter PhotoImage对象
        # image_default = image_default.subsample(2, 2)
    elif birthday_emploee_num != 0:

        for i in range(birthday_emploee_num):
            try:
                emploee_picture_array[i - 1] = 'related_files/photos/' + birthday_name_array[i - 1] + '.jpg'
                img = Image.open(emploee_picture_array[i - 1])
                resized_img = img.resize((200, 275))
                image[i - 1] = ImageTk.PhotoImage(resized_img)
                # image[i - 1] = ImageTk.PhotoImage(file= emploee_picture_array[i - 1])
                # image[i - 1] = image[i - 1].thumbnail((200, 200))
                # image[i - 1] = image[i - 1].resize((200, 200))
                # image[i - 1] = Image.open(emploee_picture_array[i - 1])
                # image[i - 1] = image[i - 1].subsample(2, 2)  # 缩放
            except Exception as e:
                print('%s缺少照片：:%s' % (birthday_name_array[i - 1], e))

                image_png = 'related_files/photos/default_image.jpg'
                image[i - 1] = PhotoImage(file=image_png)
                image[i - 1] = image[i - 1].subsample(2, 2)  # 缩放




df = pd.read_excel(r'related_files/celebration_text.xlsx')
celebration_text_array_base = df['生日祝福'].values
celebration_text_array = celebration_text_array_base
# celebration_text_num = len(celebration_text_array)
# celebration_text1 = '天天开心！天天开心！天天开心！天天开心！天天开心！天天开心！天天开心！'
# celebration_text2 = '祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！祝福！'
# celebration_text3 = '哈哈哈'

celebration_text1 = random.choice(celebration_text_array)
celebration_text2 = random.choice(celebration_text_array)
celebration_text3 = random.choice(celebration_text_array)

text_color_array = ['#04FB73','#0CF3B4','#D8FD02','#43D728','#F40B40','#886897','#00F2FF','#FFFF00','#8080FF',
                    '#F93706','#FF80C0','#80FF80','#00FF80','#80FFFF','#FF8080','#FF80FF','#FFFFFF','#FFFFFF',
                    '#FC83D1','#FDD882','#DCFE81','#A9FE81','#81FEEB','#80CFFF','#FE81D2','#FFFFFF','#FFFFFF']
text_color1 = random.choice(text_color_array)
text_color2 = random.choice(text_color_array)
text_color3 = random.choice(text_color_array)

text_font_array = ['微软雅黑 20 bold','微软雅黑 25 bold','微软雅黑 30 bold']
text1_font = random.choice(text_font_array)
text2_font = random.choice(text_font_array)
text3_font = random.choice(text_font_array)

text_move_speed_array = [1,3,5,7,9]
text1_move_speed = random.choice(text_move_speed_array)
text2_move_speed = random.choice(text_move_speed_array)
text3_move_speed = random.choice(text_move_speed_array)
# Move_speed = 5  # 文本移动速度
image_move_speed = 4

text_finnal_position_array = [100,150,200,250,300]
text1_finnal_position = random.choice(text_finnal_position_array)
text2_finnal_position = random.choice(text_finnal_position_array)
text3_finnal_position = random.choice(text_finnal_position_array)


def on_resize(evt):  # 维持窗口透明背景
    tk.configure(width=evt.width, height=evt.height)
    #canvas.create_rectangle(0, 0, canvas.winfo_width(), canvas.winfo_height(), fill=TRANSCOLOUR, outline=TRANSCOLOUR)
    canvas.create_rectangle(0, 0, canvas.winfo_width(), canvas.winfo_height(), fill=TRANSCOLOUR, outline=TRANSCOLOUR)

time_count = 0
time_count_10s = 0

def get_file_contents(file_name):
    #print("get_file_contents:" + file_name)
    with open(file_name, 'r', encoding='utf-8') as f:
        return f.read()
    
#从文本文件中更新祝福语到祝福语队列
def update_celebration_text_array():
    global celebration_text_array

    celebration_text_array = celebration_text_array_base.copy()
    reord_file_content = get_file_contents(r'record/record_file.txt')
    reord_file_content_arry = reord_file_content.split('\n')
    reord_file_content_arry.pop()
    text_arry_size_old = celebration_text_array.shape[0]
    #print(text_arry_size_old)
    celebration_text_array.resize(text_arry_size_old + len(reord_file_content_arry),) 
    celebration_text_array[text_arry_size_old:]  = reord_file_content_arry
    

def task():  # 每隔3ms移动弹幕文本位置
    global text1_X_axis,text2_X_axis,text3_X_axis,image_X_axis
    global celebration_text1,celebration_text2,celebration_text3
    global text_color1,text_color2,text_color3
    global text1_font,text2_font,text3_font
    global image_move_speed,text1_move_speed,text2_move_speed,text3_move_speed
    global text1_finnal_position,text2_finnal_position,text3_finnal_position
    global name_all
    global time_count
    global time_count_10s

    #global image_default
    if time_count_10s < 1000:
        time_count_10s += 1
    else:
        time_count_10s = 0
        update_celebration_text_array()

    if time_count < 100000:
        time_count = time_count + 1
    else:
        time_count = 0
        # 读取当天时间 间隔约26分钟
        today = datetime.date.today()
        today_str = today.strftime("%m-%d")  # 转换成字符串格式

        today_birthday_people()
        vocation_birthday_people()

        load_pictures()



    canvas.delete("move_text")  # 清除文本
    canvas.delete("mytext")
    canvas.delete("move_image")

    #if name_all == ''

    #canvas.create_text(550, 0, text=Name_text, tag="mytext", font="楷体 60 bold", fill="#F5A30A", anchor='n')
    canvas.create_text(550, 0, text=Name_text, tag="mytext", font="楷体 60 bold", fill="#FDD02F", anchor='n')

    # canvas.create_text(0, 0, text=Move_text, tag="mytext", font="微软雅黑 30 bold", fill="white", anchor='nw')  # 在新位置重新绘制文本
    text1_tag = canvas.create_text(text1_X_axis, 400, text=celebration_text1, tag="move_text", font=text1_font,
                                   fill=text_color1, anchor='nw')
    text2_tag = canvas.create_text(text2_X_axis, 500, text=celebration_text2, tag="move_text", font=text2_font, fill=text_color2, anchor='nw')  #调整字幕的位置和颜色
    text3_tag = canvas.create_text(text3_X_axis, 600, text=celebration_text3, tag="move_text", font=text3_font, fill=text_color3, anchor='nw')

    text1_bbox = canvas.bbox(text1_tag)  #使用canvas.bbox(tag)方法来获取文本的边界框
    text1_width = text1_bbox[2] - text1_bbox[0]  # 通过计算边界框的宽度来获取文本的像素长度。

    text2_bbox = canvas.bbox(text2_tag)  #使用canvas.bbox(tag)方法来获取文本的边界框
    text2_width = text2_bbox[2] - text2_bbox[0]  # 通过计算边界框的宽度来获取文本的像素长度。

    text3_bbox = canvas.bbox(text3_tag)  #使用canvas.bbox(tag)方法来获取文本的边界框
    text3_width = text3_bbox[2] - text3_bbox[0]  # 通过计算边界框的宽度来获取文本的像素长度。

    if birthday_emploee_num == 0 :
        # 在画布上创建图像对象
        image_id = canvas.create_image(550, 100, image=image_default, tag="move_image", anchor='n')
        #image_X_axis -= image_move_speed  # 移动位置
        # if image_X_axis < (-Screen_Width - 400):  # 字幕在某时间消失
        #     image_X_axis = Screen_Width
    elif birthday_emploee_num != 0 :
        for i in range(birthday_emploee_num):
            image_id = canvas.create_image(image_X_axis+(i-1)*300, 100, image=image[i-1], tag="move_image", anchor='nw')
        image_X_axis -= image_move_speed  # 移动位置
        if image_X_axis < (-Screen_Width+300):  # 字幕在某时间消失
            image_X_axis = Screen_Width
    # 缩放图像对象
    #scale(image_id, 0, 0, 0.5, 0.5)

    # canvas.move("move_image", -6, 0)
    text1_X_axis -= text1_move_speed  # 移动位置
    if text1_X_axis < (-text1_width-text1_finnal_position):   #字幕在某时间消失
        text1_X_axis = Screen_Width
        celebration_text1 = random.choice(celebration_text_array)
        text_color1 = random.choice(text_color_array)
        text1_font = random.choice(text_font_array)
        text1_move_speed = random.choice(text_move_speed_array)
        text1_finnal_position = random.choice(text_finnal_position_array)

    text2_X_axis -= text2_move_speed  # 移动位置
    if text2_X_axis < (-text2_width-text2_finnal_position):  # 字幕在某时间消失
        text2_X_axis = Screen_Width
        celebration_text2 = random.choice(celebration_text_array)
        text_color2 = random.choice(text_color_array)
        text2_font = random.choice(text_font_array)
        text2_move_speed = random.choice(text_move_speed_array)
        text2_finnal_position = random.choice(text_finnal_position_array)

    text3_X_axis -= text3_move_speed  # 移动位置
    if text3_X_axis < (-text3_width-text3_finnal_position):  # 字幕在某时间消失
        text3_X_axis = Screen_Width
        celebration_text3 = random.choice(celebration_text_array)
        text_color3 = random.choice(text_color_array)
        text3_font = random.choice(text_font_array)
        text3_move_speed = random.choice(text_move_speed_array)
        text3_finnal_position = random.choice(text_finnal_position_array)

    tk.after(10, task)


def on_drag(event):
    global start_x, start_y
    x = tk.winfo_pointerx()
    y = tk.winfo_pointery()
    dx = x - start_x
    dy = y - start_y
    tk.geometry(f"+{tk.winfo_x() + dx}+{tk.winfo_y() + dy}")
    start_x = x
    start_y = y


def on_press(event):
    global start_x, start_y
    start_x = tk.winfo_pointerx()
    start_y = tk.winfo_pointery()

def on_release(event):
    pass

tk = Tk()
Screen_Width = tk.winfo_screenwidth()  # 获取屏幕横向分辨率
Screem_Height = tk.winfo_screenheight()  #获取屏幕纵向分辨率
text1_X_axis = text2_X_axis = text3_X_axis = image_X_axis = Screen_Width
#tk.maxsize(Screen_Width + 10, 100)  # 限制窗口最大宽度为屏幕分辨率宽度, tkinter好像不能重新设置已经创建的窗口的大小
tk.maxsize(Screen_Width, Screem_Height)  #限制窗口的最大尺寸
#tk.geometry('2560x100+-5+-15')  # 向屏幕外移动一些像素隐藏窗口上和左边框
#tk.geometry('2560x100')
#tk.geometry('1000x1000+200+100')  #调整窗口的尺寸和位置
tk.geometry('1100x600')
# tk.geometry('500x300')
tk.title('生日祝福')

tk.overrideredirect(True)  # 无标题栏
#tk.Toplevel(True)
tk.attributes('-topmost', True)  # 置顶
tk.resizable(False, False)  # 窗口大小不可变
# tk.resizable(True, True)

TRANSCOLOUR = '#010358'  # 设置某个颜色为背景色，后续把此颜色设置为透明色，使窗口透明。设置窗口背景色为灰色  //设为black时不知道为什么拖不动窗口
#TRANSCOLOUR = '#F1F1F1'

#TRANSCOLOUR = 'black'
tk.wm_attributes('-transparentcolor', TRANSCOLOUR)  # 设置灰色为透明色，这样所有灰色区域都被认为是透明的了

canvas = Canvas(tk)  # 创建画布

canvas.pack(fill=BOTH, expand=Y)
#canvas.relief = "raised"

# # 绑定鼠标事件
canvas.bind('<B1-Motion>', on_drag)
canvas.bind('<ButtonPress-1>', on_press)
canvas.bind('<ButtonRelease-1>', on_release)


load_pictures()

# # 加载图片
# #image = PhotoImage(file="image.png")
# emploee_picture_array = birthday_name_array = name_all.split('  ')
#
# #length = len(birthday_name_array)
#
# image = []
# for i in range(birthday_emploee_num) :
#     image.append(0)
#
# if birthday_emploee_num == 0 :
#     image_png = 'related_files/photos/default_image.jpg'
#     image_default = Image.open(image_png)
#     resized_img = image_default.resize((700, 275))
#     image_default = ImageTk.PhotoImage(resized_img)
#     #image_default = PhotoImage(file=image_png)
#     # 将缩放后的图片转换为Tkinter PhotoImage对象
#     #image_default = image_default.subsample(2, 2)
# elif birthday_emploee_num != 0 :
#
#     for i in range(birthday_emploee_num):
#         try:
#             emploee_picture_array[i - 1] = 'related_files/photos/' + birthday_name_array[i - 1] + '.jpg'
#             img = Image.open( emploee_picture_array[i - 1])
#             resized_img = img.resize((200, 275))
#             image[i - 1] = ImageTk.PhotoImage(resized_img)
#             # image[i - 1] = ImageTk.PhotoImage(file= emploee_picture_array[i - 1])
#             # image[i - 1] = image[i - 1].thumbnail((200, 200))
#             # image[i - 1] = image[i - 1].resize((200, 200))
#             # image[i - 1] = Image.open(emploee_picture_array[i - 1])
#             # image[i - 1] = image[i - 1].subsample(2, 2)  # 缩放
#         except Exception as e:
#             print('%s缺少照片：:%s' %(birthday_name_array[i - 1], e))
#
#             image_png = 'related_files/photos/default_image.jpg'
#             image[i - 1] = PhotoImage(file=image_png)
#             image[i - 1] = image[i - 1].subsample(2, 2)  # 缩放



#image = PhotoImage(file="image.png")

tk.bind('<Configure>', on_resize)  # tkinter会自动重绘窗口导致透明背景失效, 重写事件维持透明背景
update_celebration_text_array()
tk.after(10, task)  # 递归循环,每隔10ms移动弹幕文本位置


tk.mainloop()




# # 导入Tkinter模块
# import tkinter as tk
# # 创建Tkinter窗口实例
# root = tk.Tk()
# # 定义字幕文本
# text = "欢迎来到Python世界！"
# # 创建一个Canvas实例
# canvas = tk.Canvas(root, width=400, height=50)
# # 添加Canvas实例
# canvas.pack()
# # 计算文本长度
# text_len = len(text) * 10
# # 设置初始坐标
# x = canvas.winfo_width()
# y = canvas.winfo_height() / 2
# # 创建滚动字幕函数
# def scroll_text():
#     global x
#     # 移动文本
#     canvas.move("text", -1, 0)
#     x -= 1
#     # 如果文本越界，返回到右侧
#     if x< -text_len:
#     x = canvas.winfo_width()
#     canvas.move("text", x, 0)
#     # 重复调用滚动函数
#     canvas.after(10, scroll_text)
# # 添加滚动字幕文本
# canvas.create_text(x, y, text=text, fill="white", font=("Helvetica", 20), tags=("text",))
# # 调用滚动函数
# scroll_text()
# # 运行窗口
# root.mainloop()
