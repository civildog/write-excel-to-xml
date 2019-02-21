import win32com.client
import os
import shutil
import time
import pythoncom
import _thread
import xlsxwriter

root_path = [r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\2 路面检表\施工检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表\施工记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\2 路面检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\2 路面检表\监理检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\3 路面外观记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\4 路面记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\4 路面记录表\监理记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表\检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\2 交安工程监理单位用表\2 交安检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\2 交安工程监理单位用表\2 交安检表\检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表\检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\2 绿化工程用表（监理单位使用）\2 绿化检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\2 绿化工程用表（监理单位使用）\2 绿化检表\检表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\2 声屏障工程用表（监理单位使用）\3 声屏障外观鉴定记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表\外观记录表",
             r"C:\Users\admin\Desktop\替换筑业\第三批\8 声屏障工程（第八册）\2 声屏障工程用表（监理单位使用）\3 声屏障外观鉴定记录表\外观记录表",
             ]


def word_to_excel(root_dir_lists):

    """将word（doc）格式转化为excel（xls）"""

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = 1
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    for root_dir in root_dir_lists:
        lists_dir = os.listdir(root_dir)
        for list_dir in lists_dir:
            addition = os.path.splitext(list_dir)
            if addition[1] == ".doc":
                dir_item = os.path.join(root_dir,list_dir)
                doc = word.Documents.Open(dir_item)
                name = doc.Name.rstrip(".doc")
                # print(name)
                new_path = os.path.join(root_dir, name+".htm")
                doc.SaveAs(new_path, 8) # 另存为网页文件
                doc.Close() # 关闭word文档

                # 另存为excel文件
                try:
                    work = excel.WorkBooks.Open(new_path)
                    excel_path = os.path.join(root_dir, name+".xls")
                    work.SaveAs(excel_path,1)
                    work.Close()
                    # new_work = excel.Workbooks.Open(excel_path)
                    # new_work.SaveAs(os.path.join(root_dir,name+".xlsx"))
                except PermissionError:
                    continue
    word.Quit()
    excel.Quit()

    # 将由于保存htm而产生的多余文件夹删除
    for root_dir in root_dir_lists:
        check_lists = os.listdir(root_dir)
        for item in check_lists:
            add_lists = os.path.splitext(item)
            if add_lists[1] == ".files":
                shutil.rmtree(os.path.join(root_dir, item))
            elif add_lists[1] == ".htm":
                os.remove(os.path.join(root_dir, item))
            elif add_lists[1] == ".doc":
                os.remove(os.path.join(root_dir, item))
            elif add_lists[1] == ".docx":
                os.remove(os.path.join(root_dir, item))


# word_to_excel(root_path)

add_list = [r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\3 路面外观记录表\施工外观记录表",
            r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\3 路面外观记录表\监理外观记录表",
            ]


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
            except:
                print("无法找到指定内容！")
        else:
            print("不是文件！")
    else:
        print("目录不存在！")
    return adr_list


def xls_to_xlsx(root_dir_lists):
    """将xls格式转化为xlsx"""
    excel_app = win32com.client.gencache.EnsureDispatch("Excel.Application")
    for format_path in root_dir_lists:
        format_path_lists = os.listdir(format_path)
        for format_path_list in format_path_lists:
            chart_path = os.path.join(format_path, format_path_list)
            if os.path.splitext(chart_path)[1] == ".xls":
                print(chart_path)
                workbook = excel_app.Workbooks.Open(chart_path)
                # print(chart_path.rstrip(".xls"))
                path = chart_path.rstrip(".xls")+".xlsx"
                print(path)
                workbook.SaveAs(path, FileFormat=51)
                workbook.Close()
                os.remove(chart_path)
            elif os.path.splitext(chart_path)[1] == ".XLS":
                print(chart_path)
                workbook = excel_app.Workbooks.Open(chart_path)
                path = chart_path.rstrip(".XLS") + ".xlsx"
                print(path)
                workbook.SaveAs(path, FileFormat=51)
                workbook.Close()
                os.remove(chart_path)
            elif os.path.splitext(os.path.splitext(chart_path)[0])[1] == ".XLS":
                os.remove(chart_path)


def keys_ac_tab(excel_books_sheet, row_arg, col_arg, addtions, expect_key):
    """根据提供的条件来写入key值(填写特定字段的key)"""

    for num in range(col_arg-1, 0, -1):
        if addtions in ["施工单位：", "监理单位："]:
            try:
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() == addtions:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = expect_key
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 2
                    print("表头公司信息填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value == expect_key:
                    break
            except Exception:
                continue
        else:
            try:
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() == addtions:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = expect_key
                    print("表头其他信息填充成功")
                elif excel_books_sheet.Cells(row_arg, num).Value == expect_key:
                    break
            except Exception:
                continue


def keys_ac_con(excel_books_sheet, row_arg, col_arg, addtions):
    """根据固定内容写key值，唯一性"""
    for num in range(col_arg-1, 0, -1):
        try:
            if addtions == "编  号：":
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() == addtions:
                    if (col_arg > 26) and (col_arg <= 52):
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%PT" + "A{}".format(
                            chr(64 + col_arg - 26)) + str(row_arg)
                    else:
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%PT" + chr(64 + col_arg) + str(row_arg)
                    print("编号填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value[:3] == "%TC":
                    break
            elif addtions == "复选框":
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() in ["是", "否"]:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%CH"+chr(64+col_arg)+str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 2
                    print("复选框填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value[:3] == "%CH":
                    break

            elif addtions == "日期":
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() == "日期：" or str(excel_books_sheet.Cells(row_arg, num).Value).strip()[-2:] in ["日期", "时间"]:
                    if (col_arg > 26) and (col_arg <= 52):
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%DT" + "A{}".format(
                            chr(64 + col_arg - 26)) + str(row_arg)
                    else:
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%DT" + chr(64 + col_arg) + str(row_arg)
                    print("日期填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value[:3] == "%DT":
                    break
            elif addtions == "签名":
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() in ["检验负责人：", "检测：", "签名：", "检查人：", "质检负责人：", "监理员：",
                                                                                "质检人：", "专业监理工程师：", "记录：", "复核：",
                                                                                "总监理工程师："] or\
                        ((excel_books_sheet.Cells(row_arg, num).Value is not None) and
                         (excel_books_sheet.Cells(row_arg, num).Value[-3:] in ["签名："])):
                    if (col_arg > 26) and (col_arg <= 52):
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%PT" + "A{}".format(
                            chr(64 + col_arg - 26)) + str(row_arg)
                    else:
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%PT" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 2
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    print("签名填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value[:3] == "%PT":
                    break
            elif addtions == "页码":
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() == "第":
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%CP"+chr(64+col_arg)+str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    print("当前页填充成功！")
                elif str(excel_books_sheet.Cells(row_arg, num).Value).strip() in ["页， 共", "页，共"]:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%AP" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    print("总页码填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value[:3] in ["%CP", "%AP"]:
                    break
            elif addtions == "其他":
                if str(excel_books_sheet.Cells(row_arg, num).Value).strip() in ["检查总尺数", "合格尺数", "合格率（%）", "起讫桩号：",
                                                                   "检测段落", "规范要求", "结构类型：", "铺筑桩号：",
                                                                   "天气：", "气温：", "备注", "天气情况", "最高", "最低",
                                                                   "试件编号", "允许偏差（mm）", "纵缝顺直度允许偏差（mm）",
                                                                   "横缝顺直度允许偏差（mm）", "拍摄人", "拍摄需显示的人物、工程部位或背景內容",
                                                                   "钢筋规格型号", "合同段", "质量保证资料"]:
                    if (col_arg > 26) and (col_arg <= 52):
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + "A{}".format(
                            chr(64 + col_arg - 26)) + str(row_arg)
                    else:
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    print("其他信息填充成功！")
                elif excel_books_sheet.Cells(row_arg, num).Value[:3] == "%TC":
                    break
        except Exception:
            continue


def check_row(excel_books_sheet, row_arg, addtions):
    """判断符合条件的行"""
    contents = list(list(excel_books_sheet.UsedRange.Rows(row_arg).Value)[0])
    content = []
    for i in range(len(contents)):
        if contents[i] is not None:
            content.append(str(contents[i]))
    con_str = "".join(content)
    if addtions in con_str:
        return True
    else:
        return False


def check_col(excel_books_sheet, col_arg, addtions):
    """判断符合条件的列"""
    contents = list(list(excel_books_sheet.UsedRange.Columns(col_arg).Value)[0])
    content = []
    for i in range(len(contents)):
        if contents[i] is not None:
            content.append(str(contents[i]))
    con_str = "".join(content)
    if addtions in con_str:
        return True
    else:
        return False


def keys_ac_col(name, excel_books_sheet, row_arg, col_arg):
    """根据列填充key值"""
    check_list = ["长度", "宽度", "厚度", "高度"]
    if ("检表" in name) or ("检验表" in name):
        if col_arg >= 2:
            if (str(excel_books_sheet.Cells(row_arg, col_arg-1).Value).strip() == "（") and (str(excel_books_sheet.Cells(row_arg, col_arg+1).Value).strip() == "）"):
                excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                print("规范值、设计值、允许值填充成功！")

        for num in range(row_arg-1, 0, -1):
            try:
                if str(excel_books_sheet.Cells(num, col_arg).Value).strip() in ["实测值或实测偏差值", "实测值或偏差值",]:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%EP"+chr(64+col_arg)+str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("实测值或实测偏差值填充成功！")
                    break
                    # elif excel_books_sheet.Cells(num, col_arg).Value[:3] == "%TC":
                    #     break
                elif str(excel_books_sheet.Cells(num, col_arg).Value).strip() in ["满足设计要求", "不小于设计值", "≤设计验收弯沉值",
                                                                                  "在合格标准内", "≥设计值",
                                                                                  ]:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("Through")
                    break
                elif "不小于设计值，设计未规定时不小于" in str(excel_books_sheet.Cells(num, col_arg).Value).strip():
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("Through")
                    break
            except Exception:
                continue
    elif "评定表" in name:
        if col_arg >= 2:
            if (str(excel_books_sheet.Cells(row_arg, col_arg-1).Value).strip() in ["(", "（"]) or (str(excel_books_sheet.Cells(row_arg, col_arg+1).Value).strip() in [")","）"]):
                if (col_arg > 26) and (col_arg <= 52):
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + "A{}".format(chr(64 + col_arg-26)) + str(row_arg)
                else:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                print("规范值、设计值、允许值填充成功！")

        for num in range(row_arg-1, 0, -1):
            try:
                if str(excel_books_sheet.Cells(num, col_arg).Value).strip() in ["平均值、{}代表值".format("\n"), "合格率{}（%）".format("\n"), "合格判定"]:
                    if (col_arg > 26) and (col_arg <= 52):
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + "A{}".format(
                            chr(64 + col_arg - 26)) + str(row_arg)
                    else:
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("合格判定信息填充成功！")
                    break
                elif str(excel_books_sheet.Cells(num, col_arg).Value).strip() in ["1.0", "2.0", "3.0", "4.0", "5.0",
                                                                                  "6.0", "7.0", "8.0", "9.0", "10.0",
                                                                                  ]:
                    if (col_arg > 26) and (col_arg <= 52):
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%EP" + "A{}".format(
                            chr(64 + col_arg - 26)) + str(row_arg)
                    else:
                        excel_books_sheet.Cells(row_arg, col_arg).Value = "%EP" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("表单内容信息填充成功！")
                    break

            except Exception:
                continue
    elif ("记录表" in name) and ("鉴定" not in name):
        for num in range(row_arg - 1, 0, -1):
            try:
                if ("图" in str(excel_books_sheet.Cells(num, col_arg).Value)) or\
                    ("照片" in str(excel_books_sheet.Cells(num, col_arg).Value)) or \
                        str(excel_books_sheet.Cells(num, col_arg).Value).strip() == "图示和备注：":
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%PT" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    print("图片填充成功！")
                    break

                elif str(excel_books_sheet.Cells(num, col_arg).Value).strip() in ["桩      号", "实测厚度（mm）", "设计厚度（mm）", "偏差（mm）", "备注",
                                                                   "桩号位置", "检查项目和内容", "设计或规定值", "检查结果记录", "检  查  意  见",
                                                                   1,1.0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 10.0, "1", "10","结构层名称", "施工段落", "取样位置",
                                                                   "芯样描述", "试件编号", "序号", "差值（cm）", "施工桩号", "≥145", "≥135",
                                                                   "≥130", "≥70", "桩号及部位", "φ200mmUPVC管长(cm)", "设计值", "偏差",
                                                                   "管长(cm)", "桩号", "设计值或允许偏差", "检查情况记录", "检查桩号或部位",
                                                                                ]:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("其他信息填充成功")
                    break

                elif str(excel_books_sheet.Cells(num, col_arg).Value).strip() in ["实测值", "实测值或实测偏差值", "实测值（mm）"]:
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%MP" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("实测值填充成功！")
                    break

                elif "日期" in str(excel_books_sheet.Cells(num, col_arg).Value):
                    excel_books_sheet.Cells(row_arg, col_arg).Value = "%DT" + chr(64 + col_arg) + str(row_arg)
                    excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                    excel_books_sheet.Cells(row_arg, col_arg).Font.Name = "宋体"
                    excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                    print("日期填充成功！")
                    break

                else:
                    for item in check_list:
                        if item in str(excel_books_sheet.Cells(num, col_arg).Value):
                            excel_books_sheet.Cells(row_arg, col_arg).Value = "%TC" + chr(64 + col_arg) + str(row_arg)
                            excel_books_sheet.Cells(row_arg, col_arg).Font.ColorIndex = 1
                            excel_books_sheet.Cells(row_arg, col_arg).HorizontalAlignment = 3
                            print("长度、厚度、高度、宽度填充成功！")
                            break
            except Exception:
                continue


def write_key_thread(path):
    """创建一个线程"""
    if os.path.isdir(path):
        files_list = os.listdir(path)
        for file in files_list:
            file_path = os.path.join(path, file)
            if (os.path.isfile(file_path)) and (os.path.splitext(file_path)[1] == ".xlsx"):
                time.sleep(5)
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = 1
                excel.DisplayAlerts = True
                work = excel.Workbooks.Open(file_path)
                print(work.Name)
                sheet = work.Worksheets(1)
                info = sheet.UsedRange
                info.Font.Name = "宋体"
                print("修改字体成功！")
                nrows = info.Rows.Count
                ncols = info.Columns.Count

                if ("检表" in work.Name) or ("检验表" in work.Name):
                    """填充检表的keys值"""
                    for row in range(1, nrows + 1):
                        add = check_row(sheet, row, "检验表")
                        ADD = check_row(sheet, row, "记录表")
                        if add and ADD:
                            sheet.Range(sheet.Rows(row + 1), sheet.Rows(nrows)).Font.Size = 9
                            print("修改字号成功！")
                            break

                    for row in range(1, nrows+1):
                        for col in range(1, ncols+1):
                            if sheet.Cells(row, col).Value is None:
                                keys_ac_tab(sheet, row, col, "施工单位：", "%QCConstructName")
                                keys_ac_tab(sheet, row, col, "监理单位：", "%QCSurvName")
                                keys_ac_tab(sheet, row, col, "合同段：", "%QCTenderName")
                                keys_ac_con(sheet, row, col, "编  号：")
                                keys_ac_tab(sheet, row, col, "工程名称", "%QCSubitemName")
                                keys_ac_tab(sheet, row, col, "桩号及工程部位", "%QCPositionName")
                                keys_ac_con(sheet, row, col, "编  号：")
                                keys_ac_con(sheet, row, col, "复选框")
                                keys_ac_con(sheet, row, col, "日期")
                                keys_ac_con(sheet, row, col, "签名")

                    for row in range(1, nrows+1):
                        end_row = check_row(sheet, row, "意见：")
                        if not end_row:
                            for col in range(1, ncols+1):
                                if sheet.Cells(row, col).Value is None:
                                    keys_ac_col(work.Name, sheet, row, col)
                        else:
                            break

                    for row in range(nrows, 2, -1):
                        for col in range(2, ncols+1):
                            if sheet.Cells(row, col).Value is None:
                                if (check_row(sheet, row - 1, "自检是否合格")) or (check_row(sheet, row - 1, "检验是否合格")):
                                    sheet.Cells(row, col).Value = "%TC" + chr(64 + col) + str(row)
                                    sheet.Cells(row, col).Font.ColorIndex = 1
                                    sheet.Cells(row, col).HorizontalAlignment = 3
                                    print("意见栏填充成功！")

                    for row in range(1, nrows + 1):
                        end_row = check_row(sheet, row, "检验表")
                        END_ROW = check_row(sheet, row, "记录表")
                        if end_row:
                            # for col in range(2, ncols+1):
                            try:
                                sheet.Cells(row-1, 2).Value = "%QCProjectName"
                                sheet.Cells(row-1, 2).Font.Underline = False
                                sheet.Cells(row-1, 2).Font.ColorIndex = 1
                                sheet.Cells(row-1, 2).Font.Name = "宋体"
                                sheet.Cells(row-1, 2).HorizontalAlignment = 3
                                break
                            except Exception:
                                continue

                elif "评定表" in work.Name:

                    """填充评定表keys值"""

                    address = find_adr(file_path, "基本{}要求".format("\n"))
                    for row in range(1, int(address[0])):
                        for col in range(1, ncols+1):
                            if sheet.Cells(row, col).Value is None:
                                keys_ac_tab(sheet, row, col, "施工单位：", "%QCConstructName")
                                keys_ac_tab(sheet, row, col, "监理单位：", "%QCSurvName")
                                keys_ac_tab(sheet, row, col, "合同段：", "%QCTenderName")
                                keys_ac_tab(sheet, row, col, "所属建设项目（合同段）：", "%QCTenderName")
                                keys_ac_tab(sheet, row, col, "分项工程名称：", "%QCSubitemName")
                                keys_ac_tab(sheet, row, col, "所属分部工程名称：", "%QCBranchName")
                                keys_ac_tab(sheet, row, col, "所属分部工程名称：", "%QCBranchName")
                                keys_ac_tab(sheet, row, col, "所属单位工程：", "%QCCompanyName")
                                keys_ac_tab(sheet, row, col, "分项工程编号：", "%QCSubitemNumber")
                                keys_ac_tab(sheet, row, col, "工程名称", "%QCSubitemName")
                                keys_ac_tab(sheet, row, col, "桩号及工程部位", "%QCPositionName")
                                keys_ac_tab(sheet, row, col, "工程部位：", "%QCPositionName")
                                keys_ac_con(sheet, row, col, "编  号：")

                    # for row in range(1, nrows + 1):
                    #     add = check_row(sheet, row, "评定表")
                    #     if add:
                    #         sheet.Range(sheet.Rows(row + 1), sheet.Rows(nrows)).Font.Size = 9
                    #         print("修改字号成功！")
                    #         break

                    for row in range(1, nrows+1):
                        for col in range(1, ncols+1):
                            if sheet.Cells(row, col).Value is None:
                                keys_ac_con(sheet, row, col, "复选框")
                                keys_ac_con(sheet, row, col, "日期")
                                keys_ac_con(sheet, row, col, "签名")
                                keys_ac_con(sheet, row, col, "其他")
                            elif sheet.Cells(row, col).Value == "年    月    日":
                                if (col > 26) and (col <= 52):
                                    sheet.Cells(row, col).Value = "%DT" + "A{}".format(chr(64 + col-26)) + str(row)
                                else:
                                    sheet.Cells(row, col).Value = "%DT" + chr(64 + col) + str(row)
                            elif sheet.Cells(row, col).Value == "□":
                                if (col > 26) and (col <= 52):
                                    sheet.Cells(row, col).Value = "%CH" + "A{}".format(chr(64 + col-26)) + str(row)
                                else:
                                    sheet.Cells(row, col).Value = "%CH" + chr(64 + col) + str(row)

                    for row in range(1, nrows+1):
                        end_row = check_row(sheet, row, "质量保证资料")
                        if not end_row:
                            for col in range(1, ncols+1):
                                if sheet.Cells(row, col).Value is None:
                                    keys_ac_col(work.Name, sheet, row, col)
                        else:
                            break

                    for row in range(1, nrows + 1):
                        end_row = check_row(sheet, row, "评定表")
                        if end_row:
                            try:
                                sheet.Cells(row-1, 2).Value = "%QCProjectName"
                                sheet.Cells(row-1, 2).Font.Underline = False
                                sheet.Cells(row-1, 2).Font.ColorIndex = 1
                                sheet.Cells(row-1, 2).Font.Name = "宋体"
                                sheet.Cells(row-1, 2).HorizontalAlignment = 3

                                sheet.Range(sheet.Rows(row + 1), sheet.Rows(nrows)).Font.Size = 9  # 修改正文表格的字号
                                print("修改字号成功！")
                                break
                            except Exception:
                                continue
                    # for row in range(1, nrows+1):
                    #     for col in range(1, ncols+1):
                    #         if sheet.Cells(row, col).Value == "年    月    日":
                    #             if (col > 26) and (col <= 52):
                    #                 sheet.Cells(row, col).Value = "%DT" + "A{}".format(chr(64 + col-26)) + str(row)
                    #             else:
                    #                 sheet.Cells(row, col).Value = "%DT" + chr(64 + col) + str(row)
                    #         elif sheet.Cells(row, col).Value == "□":
                    #             if (col > 26) and (col <= 52):
                    #                 sheet.Cells(row, col).Value = "%CH" + "A{}".format(chr(64 + col-26)) + str(row)
                    #             else:
                    #                 sheet.Cells(row, col).Value = "%CH" + chr(64 + col) + str(row)

                elif ("记录表" in work.Name) and ("鉴定" not in work.Name):

                    # for row in range(1, nrows + 1):
                    #     add = check_row(sheet, row, "自检")
                    #     if add:
                    #         sheet.Range(sheet.Rows(row + 1), sheet.Rows(nrows)).Font.Size = 9
                    #         break

                    for row in range(1, nrows + 1):
                        end_row = check_row(sheet, row, "记录表")
                        if end_row:
                            try:
                                sheet.Cells(row-1, 2).Value = "%QCProjectName"
                                sheet.Cells(row-1, 2).Font.Underline = False
                                sheet.Cells(row-1, 2).Font.ColorIndex = 1
                                sheet.Cells(row-1, 2).Font.Name = "宋体"
                                sheet.Cells(row-1, 2).HorizontalAlignment = 3

                                sheet.Range(sheet.Rows(row + 1), sheet.Rows(nrows)).Font.Size = 9  # 调整正文表格部分字号
                                print("修改字号成功！")
                                break
                                # break
                            except Exception:
                                continue
                    for row in range(1, nrows+1):
                        end_row_1 = check_row(sheet, row, "检查人：")
                        end_row_2 = check_row(sheet, row, "检查总尺数")
                        end_row_3 = check_row(sheet, row, "质检人：")
                        end_row_4 = check_row(sheet, row, "备注")
                        end_row_5 = check_row(sheet, row, "天气情况")
                        if (not end_row_1) and (not end_row_2) and (not end_row_3) and (not end_row_4) and (not end_row_5):
                            for col in range(1, ncols+1):
                                if sheet.Cells(row, col).Value is None:
                                    keys_ac_col(work.Name, sheet, row, col)
                        else:
                            break
                    for row in range(1, nrows+1):
                        for col in range(1, ncols+1):
                            if sheet.Cells(row, col).Value is None:
                                keys_ac_tab(sheet, row, col, "合同段：", "%QCTenderName")
                                keys_ac_tab(sheet, row, col, "施工单位：", "%QCConstructName")
                                keys_ac_tab(sheet, row, col, "监理单位：", "%QCSurvName")
                                keys_ac_con(sheet, row, col, "编  号：")
                                keys_ac_tab(sheet, row, col, "工程名称", "%QCSubitemName")
                                keys_ac_tab(sheet, row, col, "工程名称：", "%QCSubitemName")
                                keys_ac_tab(sheet, row, col, "桩号及工程部位", "%QCPositionName")
                                keys_ac_tab(sheet, row, col, "工程部位", "%QCPositionName")
                                keys_ac_tab(sheet, row, col, "桩号部位", "%QCPositionName")
                                keys_ac_tab(sheet, row, col, "桩号及部位", "%QCPositionName")
                                keys_ac_tab(sheet, row, col, "桩    号", "%QCPositionName")
                                keys_ac_con(sheet, row, col, "复选框")
                                keys_ac_con(sheet, row, col, "日期")
                                keys_ac_con(sheet, row, col, "签名")
                                keys_ac_con(sheet, row, col, "页码")
                                keys_ac_con(sheet, row, col, "其他")

                work.Close(-1)
                excel.Quit()
            elif os.path.isdir(file_path):
                write_key_thread(file_path)


# if __name__ == "__main__":
#     write_tab_path = [r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\2 路面检表\施工检表",
#                       r"C:\Users\admin\Desktop\替换筑业\第三批\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表\施工记录表",
#                       r"C:\Users\admin\Desktop\替换筑业\第三批\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表\检表",
#                       r"C:\Users\admin\Desktop\替换筑业\第三批\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表\检表",
#                       ]
#     for path_item in write_tab_path[1]:
#         write_key_thread(path_item)
#         if path_item == write_tab_path[-1]:
#             print("keys值填写完成！恭喜你，可以进入到下一步操作！祝您猪年大吉大利！天天吃鸡！")
