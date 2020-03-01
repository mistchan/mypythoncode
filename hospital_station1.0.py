#!D:\Python37
# -*- coding: utf-8 -*-
# author:mistchan
# 本程序运行需要依赖的目录有3个：配置文件夹：'.\wsetup'; 存放有目标药品截图的文件夹:'.\targartimage'; 存放有细胞图片的文件夹:'.\cellimages'
import glob
import sys
import os
import re
import shelve
import shutil
import time
import tkinter
import tkinter.filedialog
import tkinter.messagebox
import winsound
import pandas as pd
import pyautogui
import win32com.client
import multiprocessing as mp
from PIL import Image
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from hashlib import md5
from pwd import Pwd
# 打开配置文件
datasaved = shelve.open('.\\wsetup\\win32')
imWidth = datasaved['imWidth']
imHeight = datasaved['imHeight']
PatientIdRegion = datasaved['PatientIdRegion']
patientIdwidth = datasaved['patientIdwidth']
datasaved.close()


def remove_dir(dirs):
    # dirs为列表类型，为需要删除的目录列表
    for each in dirs:
        shutil.rmtree(each)


def make_dir(dirs):
    # dirs为列表类型，为需要生成的目录列表
    for each in dirs:
        if not os.path.isdir(each):
            os.makedirs(each)
        else:
            shutil.rmtree(each)
            os.makedirs(each)


dir_list_temp = dir_list_removed = (
    '.\\white_image_dir', '.\\blue_image_dir', '.\\results\\done', '.\\im_cropped', '.\\temp',
    '.\\result', '.\\blue_to_white_temp')


# 定义函数查找第一级文件夹下特定名称文件夹名字的路径的全部列表。
# 以第一级文件夹路径（路径格式）和最后一级文件夹名（字符串）为参数，返回符合条件的完整路径列表。
def path_fit(root_dir_path, end_dir_str):
    dir_walk = os.walk(root_dir_path)
    dir_list_all = []
    for each_dir_g in dir_walk:
        for n in each_dir_g[1]:
            if n == str(end_dir_str):
                dir_list_all.append(os.path.join(each_dir_g[0], n))
    return dir_list_all


class AskDir(object):
    dir_im = dir_name = month_to_check = year_to_check = ''
    cur_year = time.strftime("%Y", time.localtime())
    cur_month = time.strftime("%m", time.localtime())

    def __init__(self, root):
        self.root = root

        self.root.title('选择包含截图的总文件夹:')
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()

        ww = 500
        wh = 200
        self.root.geometry("%dx%d+%d+%d" % (ww, wh, (sw - ww) / 2, (sh - wh) / 2))

        self.l1 = tkinter.Label(root, text='输入需找到的截图文件所在的目录名', height=2)
        self.l1.pack()
        self.df = tkinter.StringVar()
        self.df.set('all')
        self.e1 = tkinter.Entry(root, textvariable=self.df)
        self.e1.pack()

        self.y = tkinter.StringVar()
        self.y.set(self.cur_year)

        self.m = tkinter.StringVar()
        self.m.set(self.cur_month)

        tkinter.Label(root, text='输入查询的年月').pack()

        f1 = tkinter.Frame(root)
        f1.pack()
        self.e3 = tkinter.Entry(f1, textvariable=self.y)
        self.e3.pack(side=tkinter.LEFT)
        tkinter.Label(f1, text='年', height=2).pack(side=tkinter.RIGHT)

        f = tkinter.Frame(root)
        f.pack()
        self.e2 = tkinter.Entry(f, textvariable=self.m)
        self.e2.pack(side=tkinter.LEFT)
        tkinter.Label(f, text='月', height=2).pack(side=tkinter.RIGHT)
        frame_rename = tkinter.Frame(self.root)
        frame_rename.pack()

        tkinter.Button(frame_rename, text='选择文件夹', command=self.c_rename, width=10, height=2,

                       activebackground='grey', relief='groove').pack(side=tkinter.LEFT)
        tkinter.Button(frame_rename, text='退出', command=self.q_rename, width=10, height=2,

                       activebackground='grey', relief='groove').pack(side=tkinter.RIGHT)

    def c_rename(self):
        self.root.update()
        self.dir_im = tkinter.filedialog.askdirectory()
        self.dir_name = self.e1.get()
        self.month_to_check = self.e2.get()
        self.year_to_check = self.e3.get()
        self.root.destroy()

    def q_rename(self):
        try:
            sys.exit(0)
        except SystemExit:
            tkinter.messagebox.showinfo(title='退出', message='用户退出，程序结束！')
            sys.exit(0)


def time_now():
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())


def recognize(screen_shot_im):
    n_im = 1
    for targ_im in (glob.glob('.\\targartimage\\*.png')):

        im = Image.open(targ_im)

        im_b = Image.open(screen_shot_im)
        # im_bx, im_by = im_b.size
        llist = list(
            pyautogui.locateAll(im, im_b, grayscale=True, region=(168, 160, 181, 497)))  # 1366 x 768 分辨率下 所有匹配图像的坐标生成列表
        if llist:

            for eachlist in llist:
                imr = im_b.crop(
                    (eachlist[0], eachlist[1], eachlist[0] + imWidth,
                     eachlist[1] + imHeight))  # 截取形成的信息图像，宽度为778，高度为22
                imcode = im_b.crop(PatientIdRegion)  # 截取住院号，宽度为74
                imr.paste(imcode, (imWidth - patientIdwidth, 1))  # 粘贴住院号，粘贴到信息条的最右端，横坐标为两者宽度相减：778-74

                imr.save(
                    '.\\results\\done\\' + re.split(r'\\|/', screen_shot_im)[-3] + '_' +
                    re.split(r'\\|/', screen_shot_im)[-1][:-4] + '_' + str(n_im) + '.png')

                n_im += 1


def im_to_str(tal_im_dir, q):
    target_im_opened = Image.open(tal_im_dir)

    im_str = []

    for cell_im_dir_name in [
        ('drug', (0, 0, 130, 18)),
        ('doc', (495, 0, 60, 18)),
        ('dot', (250, 0, 60, 18)),
        ('month', (580, 0, 45, 18)),
        ('day', (605, 0, 40, 18))
    ]:

        for cell_im in glob.glob(f'./cell_im_dir/{cell_im_dir_name[0]}/*.png'):
            cell_im_opened = Image.open(cell_im)
            if pyautogui.locate(cell_im_opened, target_im_opened, region=(cell_im_dir_name[1])):
                im_str.append(cell_im.split(os.sep)[-1][:-4])
                break

    # print(im_str)
    q.put(im_str)


def check_blue(each_im):
    signal = 1

    im = Image.open(each_im)

    imp = im.load()

    x_i, y_i = im.size
    for i in range(x_i):
        for j in range(y_i):
            if imp[i, j] == (255, 255, 255):
                imp[i, j] = (0, 0, 0)
            elif imp[i, j] == (0, 0, 128):
                imp[i, j] = (255, 255, 255)
    im_c = im.crop((int(x_i / 2), 0, x_i, y_i))

    im_tar = glob.glob('.\\white_image_dir\\*.png')
    for e_im in im_tar:

        im_t = Image.open(e_im)
        if pyautogui.locate(im_c, im_t):
            signal = 0
            print('发现重复图片：' + str(each_im))
            break

    if signal:
        im.save('.\\blue_to_white_temp\\' + each_im.split(os.sep)[-1])
        print(str(each_im) + '未发现与现有图片重复，已成功转存。')


def de_rep_im(dirt):
    list0 = []

    list1 = glob.glob(dirt + '\\*.png')

    # 计算每张图片的md5值，并将图片路径与其md5值整合到列表list中
    for n in range(len(list1)):
        hasho = md5()
        img = open(list1[n], 'rb')
        hasho.update(img.read())
        img.close()
        list2 = [list1[n], hasho.hexdigest()]

        list0.append(list2)

    # 两两比较md5值，若相同，则删去一张图片
    m = 0
    while m < len(list0):
        t = m + 1
        while t < len(list0):
            if list0[m][1] == list0[t][1]:
                os.remove(list0[t][0])
                del list0[t]
            else:
                t += 1
        m += 1


if __name__ == '__main__':

    root1 = tkinter.Tk()
    app = AskDir(root1)
    root1.mainloop()
    dir_list_all = path_fit(app.dir_im, app.dir_name)
    year_to_check = int(app.year_to_check)
    month_to_check = int(app.month_to_check)
    name_index_w = 0
    make_dir(dir_list_temp)
    print(time_now() + ' 开始运行...')

    print(time_now() + ' 识别所有截图生成小图片...')
    pool2 = mp.Pool(processes=11)

    for each_fit_path in dir_list_all:
        for each_screen_shot_im in glob.glob(each_fit_path + '\\*.png'):
            pool2.apply_async(recognize, (each_screen_shot_im,))

    pool2.close()
    pool2.join()

    # 根据生成的小图片开始汇总分析
    print(time_now() + ' 根据小图片颜色分类整理...')
    for eld_file in glob.glob('.\\results\\done\\*.png'):
        shutil.copyfile(eld_file,
                        '.\\result\\' + str(eld_file.split(os.sep)[-1]))
    for old_file in glob.glob('.\\result\\*.png'):
        im_judge = Image.open(old_file)
        if im_judge.getpixel((0, 0)) == (0, 0, 128):
            shutil.copyfile(old_file,
                            '.\\blue_image_dir\\' + os.path.basename(old_file).rstrip('.png') + '_blue' + str(
                                name_index_w) + '.png')
            name_index_w += 1
        else:
            shutil.copyfile(old_file,
                            '.\\white_image_dir\\' + os.path.basename(old_file).rstrip('.png') + '_white' + str(
                                name_index_w) + '.png')
            name_index_w += 1
    print(time_now() + ' 图片去重...')

    de_rep_im('blue_image_dir')
    print(time_now() + ' 图片去重已完成1/2...')
    de_rep_im('white_image_dir')
    print(time_now() + ' 图片去重已完成2/2...')
    print(time_now() + ' 交叉对比两组图片并删除重复项...')

    pool3 = mp.Pool(processes=11)
    for each_blue_im in glob.glob('.\\blue_image_dir\\*.png'):
        pool3.apply_async(check_blue, (each_blue_im,))

    pool3.close()
    pool3.join()

    print(time_now() + ' 汇总所有结果图片到white_image_dir文件夹...')
    for blue_image in glob.glob('.\\blue_to_white_temp\\*.png'):
        shutil.copy(blue_image, '.\\white_image_dir')
    print(time_now() + ' 第一阶段运行完毕...')
    speak = win32com.client.Dispatch('SAPI.SPVOICE')
    winsound.Beep(2019, 5000)
    speak.Speak('请选择是否运行cell image生成程序，或跳过程序、继续下一步分析')

    while True:
        speak.Speak('请选择下一步执行程序')
        ask_if = pyautogui.confirm(title='请选择：', text='是否运行程序补充生成小图', buttons=['运行程序', '不运行程序继续', '结束程序并储存'])
        if ask_if == '运行程序':
            box1 = (497, 0, 544, 18)  # 人名
            box2 = (250, 0, 299, 18)  # 用量
            box3 = (606, 0, 640, 18)  # 日
            box4 = (583, 0, 611, 18)  # 月

            if not os.path.isdir('.\\im_cropped'):
                os.makedirs('.\\im_cropped')
            else:
                shutil.rmtree('.\\im_cropped')
                os.makedirs('.\\im_cropped')

            n = 1
            for ima in glob.glob('.\\white_image_dir\\*.png'):
                im_cell = Image.open(ima)

                imn = im_cell.crop(box1)
                imn.save('.\\im_cropped\\name_' + str(n) + '.png')
                imd = im_cell.crop(box2)
                imd.save('.\\im_cropped\\dot_' + str(n) + '.png')
                imday = im_cell.crop(box3)
                imday.save('.\\im_cropped\\day_' + str(n) + '.png')
                imm = im_cell.crop(box4)
                imm.save('.\\im_cropped\\month_' + str(n) + '.png')
                n += 1

            de_rep_im('.\\im_cropped')
            print(time_now() + r' 已补充生成小图，储存于.\\im_cropped中，请对照查看！')
        elif ask_if == '不运行程序继续':
            print(time_now() + ' 正将图片转化为文字信息...')
            q = mp.Manager().Queue()

            pool = mp.Pool(processes=11)

            for i in glob.glob('.\\white_image_dir\\*.png'):
                pool.apply_async(im_to_str, (i, q))

            pool.close()
            pool.join()

            result_im_to_str = []
            while True:
                result_im_to_str.append(q.get())
                if q.empty():
                    break
            print(time_now() + ' 汇总信息、清洗并储存...')
            xl = pd.DataFrame(result_im_to_str)
            if os.path.exists(f'result_{year_to_check}_{month_to_check}.xlsx'):
                os.remove(f'result_{year_to_check}_{month_to_check}.xlsx')
            deft_ex = pd.DataFrame()
            deft_ex.to_excel(f'result_{year_to_check}_{month_to_check}.xlsx')

            writer = pd.ExcelWriter(f'result_{year_to_check}_{month_to_check}.xlsx')
            xl.to_excel(writer, index=False, header=['药品', '医生', '用量', '月', '日'], sheet_name='元数据')
            writer.save()
            pd1 = pd.read_excel(f'result_{year_to_check}_{month_to_check}.xlsx')

            df = pd1

            df['日期'] = '2019.' + df['月'].astype('str').str.rjust(2, '0') + '.' + df['日'].astype('str').str.rjust(2,
                                                                                                                 '0').copy()


            def conv(x):
                temp = x.split('.')
                if int(temp[1]) < 5:
                    return '2019' + temp[1] + temp[2]
                else:
                    return '2020' + temp[1] + temp[2]


            df['日期'] = pd.to_datetime(df['日期'].copy().apply(conv), format='%Y%m%d')

            dfo = df
            dfo = dfo[dfo['月'] == month_to_check]
            dfo = dfo.reset_index()
            dfo = dfo.drop('index', axis=1)
            df.loc[df['药品'].str.startswith('替吉奥'), '用量'] = df.loc[df['药品'].str.startswith('替吉奥'), '用量'].apply(
                lambda x: x // 28 if not x % 28 else (x // 36 if not x % 36 else (x // 42 if not x % 42 else x / 28)))
            df.loc[df['药品'].str.startswith('卡培他滨'), '用量'] = df.loc[df['药品'].str.startswith('卡培他滨'), '用量'].apply(
                lambda x: x // 12 if not x % 12 else x)

            df = df.reset_index()
            df = df.drop('index', axis=1)

            pivot_a = df.pivot_table(['用量'], index=['医生'], columns=['药品'], aggfunc=sum, fill_value=0)

            df.to_excel(writer, sheet_name='数据清洗')
            pivot_a.to_excel(writer, sheet_name='最终汇总')
            writer.save()
            print(time_now() + f' 结果已储存于根目录result_{year_to_check}_{month_to_check}.xlsx文件中。')
            os.startfile(f'.\\result_{year_to_check}_{month_to_check}.xlsx')

        else:
            print(time_now() + ' 已存入数据库，运行结束！')

            sql1 = f'''
                                    CREATE TABLE {year_to_check}年{month_to_check}月_统计表
                                    SELECT
                                        case when 医生 is not null then 医生 else '合计' end AS 医生,
                                        sum( CASE WHEN 药品 = '卡培他滨片(卓仑)'  THEN 用量 ELSE 0 END ) AS 卓仑,
                                        sum( CASE WHEN 药品 = '替吉奥胶囊(苏立)'  THEN 用量 ELSE 0 END ) AS 苏立, 
                                        sum( CASE WHEN 药品 = '卡培他滨片(艾滨)'  THEN 用量 ELSE 0 END ) AS 艾滨,
                                        sum( CASE WHEN 药品 = '替吉奥胶囊(艾奕)'  THEN 用量 ELSE 0 END ) AS 艾奕,
                                        sum( CASE WHEN 药品 = '卡培他滨片(首辅)'  THEN 用量 ELSE 0 END ) AS 首辅,
                                        sum( CASE WHEN 药品 = '替吉奥胶囊(维康达)'  THEN 用量 ELSE 0 END ) AS 维康达
                                    FROM
                                        drug.{year_to_check}年{month_to_check}月 
                                    GROUP BY
                                        医生 

                                    WITH ROLLUP
                                    ORDER BY
                                        卓仑 DESC,
                                        苏立 DESC


                        '''

            pwd_link = Pwd().mysql_link()

            con = create_engine(pwd_link)
            dfo.to_sql(f'{year_to_check}年{month_to_check}月_元数据', con, schema='drug', if_exists='replace')
            df.to_sql(f'{year_to_check}年{month_to_check}月', con, schema='drug', if_exists='replace')

            Session = sessionmaker(bind=con)

            s = Session()

            s.execute(sql1)
            s.commit()
            s.close()
            break
        winsound.Beep(2019, 3000)
        speak.Speak('您选择的' + str(ask_if) + '已运行完成。请查看运行结果')
