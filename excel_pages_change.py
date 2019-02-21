import win32com.client
import os


def amend(path1, path2):  # 此函数已弃用
    """遍历所有的文件夹，并对其中的文件进行修改"""
    code = os.path.isdir(path1)
    print(code)
    if code:
        list_1 = os.listdir(path1) #列出了文件夹下所有的表单
        dirlist_0 = []
        for direct in list_1:
            rootdir_2 = os.path.join(path1, direct)
            print(rootdir_2)
            dirlist_0.append(rootdir_2)
            os.mkdir(path2+"\\"+direct)
            dirlist_1 = os.listdir(path2)
            dirname = dirlist_1.pop()
            newdir = os.path.join(path2+"\\"+dirname)
            print(newdir)
            amend(rootdir_2,newdir)


def dir_copy(path_1, path_2):
    """将path_1中的文件夹复制到path_2文件夹中（可选择单层复制/多层复制）"""
    if os.path.exists(path_1) and os.path.exists(path_2):
        path_1_lists = os.listdir(path_1)
        path_2_lists = os.listdir(path_2)
        new_2 = []
        for unit_2 in path_2_lists:
            new_2.append(os.path.join(path_2, unit_2))
        for unit_1 in path_1_lists:
            unit_path = os.path.join(path_1, unit_1)
            try:
                if os.path.isdir(unit_path) and (unit_path not in new_2):
                    if "施工" in unit_1:
                        copy_path = os.path.join(path_2, unit_1)
                        os.mkdir(copy_path)
                        dir_copy(unit_path, copy_path)  # 此行代码为多层复制用，若是单层复制的话，就将其注掉
                elif os.path.isfile(unit_path):
                    continue
                elif unit_1 in new_2:
                    print("{0}已存在于{1}路径下".format(unit_path, path_2))
                    continue
            except:
                continue

    elif not os.path.exists(path_1):
        print("{}路径不存在!".format(path_1))
    elif not os.path.exists(path_2):
        print("{}路径不存在！".format(path_2))


# path_0 = r"C:\Users\admin\Desktop\pytest"
# path_1 = r"C:\Users\admin\Desktop\1"
#
# dir_copy(path_0, path_1)

add_dir_list = [r"4 桥梁工程(第四册)\1 桥梁工程\1 施工单位\3 外观鉴定检查记录表\桥梁外观鉴定检查记录表",
                r"4 桥梁工程(第四册)\1 桥梁工程\2 监理单位\3 外观鉴定检查记录表\桥梁外观鉴定检查记录表",
                r"4 桥梁工程(第四册)\1 桥梁工程\1 施工单位\4 记录表\桥梁记录表",
                r"4 桥梁工程(第四册)\1 桥梁工程\2 监理单位\4 记录表\桥梁记录表",
                r"5 隧道工程(第五册)\1 隧道工程\1 施工单位\4 记录表\隧道记录表",
                r"5 隧道工程(第五册)\1 隧道工程\2 监理单位\4 记录表\隧道记录表",
                ]
root_1 = r"C:\Users\admin\Desktop\需要修改表格"
root_2 = r"C:\Users\admin\Desktop\excel"
# add_dir = r"4 桥梁工程(第四册)\1 桥梁工程\1 施工单位\3 外观鉴定检查记录表\桥梁外观鉴定检查记录表"
'''
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = 1
for add_dir in add_dir_list:
    root_dir = root_1+"\\"+add_dir
    amend_lists = os.listdir(root_dir)

    for amend_item in amend_lists:

        xlsxPath = root_dir+"\\"+amend_item
        status = list(os.path.splitext(xlsxPath))
        if status[1]==".xlsx":

            work = excel.Workbooks.Open(xlsxPath)
            wsheet = work.Worksheets("sheet1")
            # print(wsheet.PageSetup.TopMargin)
            # print(work.Name)
            topmargin = wsheet.PageSetup.TopMargin
            value_1 = 1.7*28.35
            value_2 = 0.2*28.35
            # print(bool(topmargin==value_1))
            # print(bool(topmargin==value_2))
            if wsheet.PageSetup.PaperSize == 9 and wsheet.PageSetup.Orientation ==2:
                #当表单为A4横向，时页边距设置
                if topmargin != value_1:
                    wsheet.PageSetup.TopMargin = 2.5*28.35
                    wsheet.PageSetup.BottomMargin = 1*28.35
                    wsheet.PageSetup.LeftMargin = 1*28.35
                    wsheet.PageSetup.RightMargin = 1*28.35

                    wsheet.PageSetup.HeaderMargin = 0*28.35
                    wsheet.PageSetup.FooterMargin = 0*28.35

                    wsheet.PageSetup.CenterHorizontally = True
                else:
                    pass
            elif wsheet.PageSetup.PaperSize ==9 and wsheet.PageSetup.Orientation ==1:
                #当表单为A4纵向时，页边距的设置
                if topmargin != value_2:
                    wsheet.PageSetup.TopMargin = 1*28.35
                    wsheet.PageSetup.BottomMargin = 1*28.35
                    wsheet.PageSetup.LeftMargin = 2.5*28.35
                    wsheet.PageSetup.RightMargin = 1*28.35

                    wsheet.PageSetup.HeaderMargin = 0*28.35
                    wsheet.PageSetup.FooterMargin = 0*28.35

                    wsheet.PageSetup.CenterVertically = True
                else:
                    pass

            elif wsheet.PageSetup.PaperSize==8 and wsheet.PageSetup.Orientation ==2:
                #当表单为A3横向时，页边距设置
                if wsheet.PageSetup.TopMargin != 1.7 * 28.35:
                    wsheet.PageSetup.TopMargin = 2.5 * 28.35
                    wsheet.PageSetup.BottomMargin = 1 * 28.35
                    wsheet.PageSetup.LeftMargin = 1 * 28.35
                    wsheet.PageSetup.RightMargin = 1 * 28.35

                    wsheet.PageSetup.HeaderMargin = 0 * 28.35
                    wsheet.PageSetup.FooterMargin = 0 * 28.35
            elif wsheet.PageSetup.PaperSize==8 and wsheet.PageSetup.Orientation==1:
                #当表单为A3纵向时，页边距设置
                if wsheet.PageSetup.LeftMargin != 0.2 * 28.35:
                    wsheet.PageSetup.TopMargin = 1 * 28.35
                    wsheet.PageSetup.BottomMargin = 1 * 28.35
                    wsheet.PageSetup.LeftMargin = 2.5 * 28.35
                    wsheet.PageSetup.RightMargin = 1 * 28.35

                    wsheet.PageSetup.HeaderMargin = 0 * 28.35
                    wsheet.PageSetup.FooterMargin = 0 * 28.35

            work.SaveAs(root_2+"\\"+add_dir+"\\"+"{}".format(work.Name))
            work.Close()
            work.quit()



# print(os.path.abspath(rootdir)) #返回绝对路径
# print(os.path.basename(rootdir)) #返回文件名称
# print(os.path.commonprefix(list)) #返回多个路径中，所有path共有的最长的路径
# print(os.path.dirname(rootdir)) #返回文件路径
# print(os.path.exists(rootdir)) #若路径存在则返回True，如果路径不存在或损坏则返回False
# print(os.path.lexists(rootdir)) #无论路径存在或损坏与否都返回True
# print(os.path.expanduser(rootdir))
# print(os.path.expandvars(rootdir))
# print(os.path.getatime(rootdir)) #返回最后一次进入路径的时间
# print(os.path.getmtime(rootdir)) #返回最后一次修改路径下文件的时间
# print(os.path.getctime(rootdir))
# print(os.path.getsize(rootdir)) #返回文件大小，如果不存在就返回错误
# print(os.path.isabs(rootdir)) #判断路径是否为绝对路径
# print(os.path.isfile(rootdir)) #判断此path是否为文件
# print(os.path.isdir(rootdir)) #判断路径是否为目录
# print(os.path.islink(rootdir)) #判断路径是否为链接
# print(os.path.ismount(rootdir)) #判断路径是否为挂载点
# print(os.path.join(rootdir,'3','4')) #把目录和文件名合成为一个路径
# print(os.path.normcase(rootdir)) #转化path中的大小写和斜杠
# print(os.path.normpath(rootdir)) #规范path路径字符串形式
# print(os.path.realpath(rootdir)) #返回path的真实路径
# print(os.path.relpath(rootdir, "需要修改表格")) #从start开始计算相对路径
# # print(os.path.samefile(path1, path2))
# # print(os.path.sameopenfile(fp1, fp2))
# # os.path.samestat(s1,s2)
# print(os.path.split(rootdir)) #返回dirname和basename组成的元组
# print(os.path.splitdrive(rootdir)) #在windows系统中应用，返回驱动器名和路径名组成的元组
# print(os.path.splitext(rootdir)) #分割路径，返回路径名和文件扩展名的元组
# print(os.path.splitunc(rootdir)) #把路径分割为路径名和加载点
# print(os.path.walk(rootdir,visit))
# print(os.path.supports_unicode_filenames) #设置是否支持unicode的路径名
'''