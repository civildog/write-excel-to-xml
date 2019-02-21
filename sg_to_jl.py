import os
import win32com.client
import excel_template as keys
from excel_pages_change import dir_copy
import sys
import time

sys.setrecursionlimit(5000)


def amend_sgfile_name(path):
    """修改施工文件表名,并返回文件名"""
    jlname = ""
    if os.path.isfile(path):
        name = os.path.basename(path)
        if ("封面" in str(name)) and ("(SG)" in str(name)) :
            jlname = name.replace("施工单位", "监理单位").replace("(SG)", "(JL)")
            print("修改（封面）完成！")
        elif '(SG)' in str(name):
            jlname = name.replace("(SG)", "(JL)")
            print("修改（SG）完成！")
    elif os.path.isdir(path):
        print("接受到了一个文件夹，请传递一个文件路径！")
    elif not os.path.exists(path):
        print("非法路径！")
    return jlname


def adr_to_pos(adr):
    """将excel的绝对地址，转化为行和列"""
    adr_list = []
    if adr:
        col_row = adr.split("$")
        row = col_row[2]
        adr_list.append(row)
        col = ord(col_row[1]) - 64
        adr_list.append(col)
    else:
        print("{}不存在".format(adr))

    return adr_list


def find_adr(path, content):
    """查找存在内容content的地址，并返回[row, col]"""
    adr_list = []
    if os.path.exists(path):
        if os.path.isfile(path):
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = 0
                work = excel.Workbooks.Open(path)
                sheet = work.Worksheets(1)
                cel = sheet.UsedRange.Find(content)
                adr = cel.Address
                col_row = adr.split("$")
                row = col_row[2]
                adr_list.append(row)
                col = ord(col_row[1]) - 64
                adr_list.append(col)
                work.Close()
                excel.Quit()
            except:
                print("无法找到指定内容！")
        else:
            print("不是文件！")
    else:
        print("目录不存在！")
    return adr_list


def amend_sgfile_con(path, content, re_content):
    """修改施工表单内容为监理表单"""
    if os.path.isfile(path):
        name = os.path.basename(path)
        if ("检表" in name) or (("记录表" in name) and ("鉴定" not in name) and ("封面" not in name)) or ("评定表" in name):
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = 1
            work = excel.Workbooks.Open(path)
            sheet = work.Worksheets(1)
            cel = sheet.UsedRange.Rows.Find(content)

            if cel is not None:
                if content in ["施工自检", "自检是否合格", "检验是否合格"]:
                    adr = cel.Address
                    col_row = adr.split("$")
                    row = col_row[2]
                    if len(col_row[1]) == 2:
                        col = ord(col_row[1][1])-64+26
                    elif len(col_row[1]) == 1:
                        col = ord(col_row[1]) - 64
                    sheet.Cells(row, col).Value = re_content
                    sheet.Cells(row, col).HorizontalAlignment = 3
                    print("修改检表成功！", content + "替换为" + re_content)
                elif content in ["检查人意见：", "监理员意见："]:
                    adr = cel.Address
                    col_row = adr.split("$")
                    row = col_row[2]
                    col = ord(col_row[1]) - 64
                    sheet.Cells(row, col).Value = re_content
                    sheet.Cells(row, col).HorizontalAlignment = 2

            work.Close(-1)
            excel.Quit()


def amend_sign(path, re_sign_con, *sign_content):
    """修改文件的签名，但是要保证替换签名的唯一性！否则会导致文件内容丢失"""
    if os.path.isfile(path):
        name = os.path.basename(path)
        if ("检表" in name) and ("封面" not in name):
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = True
            excel.DisplayAlerts = 1
            work = excel.Workbooks.Open(path)
            sheet = work.Worksheets(1)
            print("修改检表：", work.Name)
            try:
                adr = find_adr(path, sign_content)
                sheet.Cells(adr[0], adr[1]).Value = re_sign_con
                sheet.Cells(adr[0], adr[1]).HorizontalAlignment = 3
                print("修改检表成功！", sign_content + ("替换为"+re_sign_con,))
                for col in range(adr[1]-1, 0, -1):
                    sheet.Cells(adr[0], col).Value = None
            except IndexError:
                print("该检表中查询内容不存在！")

            work.Close(-1)
            excel.Quit()

        elif ("外观记录表" in name) and ("封面" not in name):
            if re_sign_con == "总监理工程师：":
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = True
                excel.DisplayAlerts = 1
                work = excel.Workbooks.Open(path)
                sheet = work.Worksheets(1)
                print("修改外观记录表：", work.Name)
                try:
                    adr = find_adr(path, "监理员：")
                    sheet.Cells(adr[0], adr[1]).Value = None
                    ncols = sheet.UsedRange.Columns.Count
                    for col in range(adr[1]+1, ncols+1):
                        if sheet.Cells(adr[0], col).Value is not None:
                            if sheet.Cells(adr[0], col).Value[:3] == "%PT":
                                sheet.Cells(adr[0], col).Value = re_sign_con
                                sheet.Cells(adr[0], col).HorizontalAlignment = 3
                                print("修改外观记录表成功！", "监理员：" + "替换为" + re_sign_con)
                                for item in range(col+1, ncols):
                                    sheet.Cells(adr[0], item).Value = None
                    for col in range(1, ncols+1):
                        if sheet.Cells(adr[0], col).Value is None:
                            keys.keys_ac_con(sheet, adr[0], col, "签名")
                except IndexError:
                    print("总监理工程师无法填充！")

                work.Close(-1)
                excel.Quit()

            else:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = True
                excel.DisplayAlerts = 0
                work = excel.Workbooks.Open(path)
                sheet = work.Worksheets(1)
                print("修改外观记录表：", work.Name)
                try:
                    adr = find_adr(path, sign_content)
                    sheet.Cells(adr[0], adr[1]).Value = re_sign_con
                    sheet.Cells(adr[0], adr[1]).HorizontalAlignment = 3
                    print("修改外观记录表成功！", sign_content + ("替换为" + re_sign_con,))
                except IndexError:
                    print("该外观记录表中查询内容不存在！")

                work.Close(-1)
                excel.Quit()

        elif ("记录表" in name) and ("鉴定" not in name) and ("封面" not in name):

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = True
            excel.DisplayAlerts = 0
            work = excel.Workbooks.Open(path)
            sheet = work.Worksheets(1)
            print("修改记录表：", work.Name)
            try:
                adr = find_adr(path, sign_content)
                sheet.Cells(adr[0], adr[1]).Value = re_sign_con
                sheet.Cells(adr[0], adr[1]).HorizontalAlignment = 3
                print("修改记录表成功！", sign_content+("替换为"+re_sign_con,))
            except IndexError:
                print("该记录表中查询内容不存在！")

            work.Close(-1)
            excel.Quit()

        elif "封面" in name:
            print("本表单为封面，不作处理！")


path_lists = [r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\2 路面检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\2 路面检表\施工检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\3 路面外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\3 路面外观记录表\施工外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表\施工记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表\检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表\检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表\外观记录表",
              ]

path_lists_1 = [r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\1 交安工程施工单位用表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）",
              r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）",
              ]


def template_change(path):
    """按照流程修改，注意调用函数的顺序不能颠倒，否则会导致表单内容出错"""

    if os.path.exists(path):
        dir_list = os.listdir(path)
        for item in dir_list:
            p = os.path.join(path, item)
            if os.path.isfile(p) and (os.path.splitext(p)[1] == ".xlsx") and ("(JL)" in os.path.basename(p)):
                amend_sgfile_con(p, "施工自检", "监理抽检")
                amend_sgfile_con(p, "自检是否合格", "抽检是否合格")
                amend_sgfile_con(p, "检验是否合格", "是否合格")
                amend_sgfile_con(p, "监理员意见：", "专业监理工程师意见：")
                amend_sgfile_con(p, "检查人意见：", "监理员意见：")
                amend_sign(p, "签名：",  "质检负责人签名：")
                amend_sign(p, "签名：",  "专业监理工程师签名：")
                amend_sign(p, "总监理工程师：")
                amend_sign(p, "专业监理工程师：", "质检负责人：")
                amend_sign(p, "监理员：", "检查人：")
                amend_sign(p, "监理员：", "质检人：")
                amend_sign(p, "专业监理工程师签名：", "质检负责人签名：")
                amend_sign(p, "监理员签名：", "检查人签名：")
            elif os.path.isdir(p):
                template_change(p)
    elif not os.path.exists(path):
        print("Error Path")


def save_as_path(paths):
    """修改名称并将其保存到对应的目录下"""
    for path in paths:
        if os.path.exists(path):
            dir_list = os.listdir(path)
            for item in dir_list:
                new_path = os.path.join(path, item)
                # try:
                if os.path.isfile(new_path):
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = True
                    excel.DisplayAlerts = 0
                    work = excel.Workbooks.Open(new_path)

                    tab_name = amend_sgfile_name(new_path)
                    tab_path = path.replace("施工单位", "监理单位").replace("1", "2", 1)
                    dir_copy(path, tab_path)
                    work.SaveAs(os.path.join(tab_path, tab_name))
                    # dir_copy(path, path_2)
                    # work.SaveAs(os.path.join(path_2, tab_name))

                    work.Close(SaveChanges=False)
                    excel.Quit()
                # elif os.path.isdir(new_path):
                #     path_3 = os.path.join(path_2, item)
                #     save_as_jl(new_path, path_3)
                else:
                    print("无法处理！！")
        else:
            print("路径不存在！请检查路径拼写是否出错！")


def save_jl(path):
    """将施工的表单修改名称后保存到统一目录下，（path）为一个目录路径"""
    if os.path.isdir(path):
        dir_list = os.listdir(path)
        if dir_list:
            for dira in dir_list:
                file_path = os.path.join(path, dira)
                if os.path.isfile(file_path) and (os.path.splitext(file_path)[1] == ".xlsx") and ("(SG)" in os.path.basename(file_path)):
                    # if "(JL)" in os.path.basename(file_path):
                    #     amend_sgfile_con(file_path, "施工自检", "监理抽检")
                    # else:
                    #     continue
                    time.sleep(5)
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = 0
                    work = excel.Workbooks.Open(file_path)

                    name = os.path.join(path, dira.replace("SG", "JL"))
                    work.SaveAs(name)
                    excel.Quit()
                elif os.path.isdir(file_path):
                    save_jl(file_path)
        else:
            print("路径下为空！")
    else:
        print("{}不是一个目录路径！".format(path))


p_list = [r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表\施工记录表",
          r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表\检表",
          r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表\检表",
          r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表\外观记录表",
          ]

if __name__ == "__main__":
    for path in p_list:
        save_jl(path)
        template_change(path)

