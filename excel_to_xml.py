from xml.dom.minidom import Document
from xml.dom import minidom
from xml.dom.minidom import parse
import win32com.client
import os
from .sg_to_jl import find_adr, adr_to_pos
import time

'''
doc = Document()
people = doc.createElement("people")
doc.appendChild(people)
aperson = doc.createElement("person")
people.appendChild(aperson)
name = doc.createElement("name")
aperson.appendChild(name)
personname = doc.createTextNode("Annie")
name.appendChild(personname)
filename = "c:/users/admin/desktop/people.xml"
f = open(filename, 'w')
# print(doc.toprettyxml(indent="\t",newl="\n",encoding="utf-8"))
# # f.write(str(doc.toprettyxml(indent="\t",newl="\n",encoding="utf-8")))
doc.writexml(filename, indent="", addindent="\t",newl="\n",encoding="utf-8")
f.close()
'''


# python生成xml文件

def generateXml():
    impl = minidom.getDOMImplementation()  # 创建一个xml dom
    #  三个参数分别对应为：namespaceUPI，qualifiedName，doctype
    doc = impl.createDocument(None, None, None)

    # 创建根元素
    rootElement = doc.createElement('Python')
    # 为根元素添加10个子元素
    for pythonId in range(10):
        # 创建子元素
        childElement = doc.createElement('python')

        # 为子元素加上id属性
        childElement.setAttribute('id', str(pythonId))

        rootElement.appendChild(childElement)
        print(childElement.getUserData(1))

        # 将拼接好的根元素加入到根元素
        doc.appendChild(rootElement)

    # 打开xml文件，准备写入
    f = open("c:/users/admin/desktop/test.xml", 'w')

    # 写入文件
    doc.writexml(f, addindent='  ', newl='\n')

    # 关闭文件
    f.close()

# generateXml()

# python 读取XML文件

# 获取python节点下的所有id属性


def getTagId():

    # 获取test.xml文件对象
    doc = parse('c:/users/admin/desktop/test.xml')
    node_list = doc.getElementsByTagName("python")
    print(node_list)
    for node in doc.getElementsByTagName("python"):
        # 获取标签Id属性
        value_str = node.getAttribute("id")
        print(value_str)
    print(doc.getElementsByTagName("Python")[0].childNodes)
    for child in doc.getElementsByTagName("Python")[0].childNodes:
        if child.nodeType == doc.COMMENT_NODE:
            print(child.data)

# getTagId()

# python 3下MD5加密，由于MD5模块在python3中被移除，在python3中使用hashlib模块进行md5操作

# import hashlib
#
# # 待加密信息
# str_1 = "This is a md5 test."
#
# # 创建md5对象
# hl = hashlib.md5()
# p = "c:/users/admin/desktop/test.xml"
# # tips，此处必须声明encode，若写法为h1.update(str) 报错为：Unicode-objects must be encodeed
# # before hashing
# hl.update(str_1.encode())
# print('MD5 加密前为：'+str_1)
# print('MD5 加密后为：'+hl.hexdigest())
#
#
# def readFile(filename, lines):
#     with open(filename, 'r') as f:
#         for line in f:
#             line = line.rstrip('\n')
#             if line.startswith('//') or len(line) == 0:
#                 continue
#             lines.append(line)
#
#
# def writeXml(filename, lines, tagNames):
#     # 创建doc
#     doc = Document()
#     #创建根节点
#     root = doc.createElement(tagNames[0])
#     doc.appendChild(root)
#
#     # 记录每层节点的最新元素
#     nodes = {0: root}
#
#     for line in lines:
#         index = line.rfind(' ')
#         level = (index+1) / 4 + 1
#         line = line.lstrip(' ')
#
#         node = doc.createElement(tagNames[level])
#         node.setAttribute('name', line)
#
#         nodes[level-1].appendChild(node)
#         nodes[level] = node
#
#     with open(filename, 'w') as f:
#         f.write(doc.toprettyxml(indent='\t'))
#
# def display(lines):
#     for line in lines:
#         print(line)
#
# if __name__ == '__main__':
#     lines = []
#     filename = 'c:/users/admin/desktop/soupUI学习笔记.txt'
#     readFile(filename, lines)
#
#     tagNames = ['SectorFile', 'Sectors', 'sector', 'sector_second']
#     writeXml(filename, lines, tagNames)
#
#
# domtree = parse('movies.xml')
# collection = domtree.documentElement
# print(collection)
# if collection.hasAttribute("shelf"):
#     print("haode")
#
# # 获取所有电影标签
#
# movies = collection.getElementsByTagName("movie")

# 打印每部电影的详细信息

# for movie in movies:
#     print("*****Movie*****")
#     if movie.hasAttribute("title"):
#         print("haha")
#
#     type = movie.getElementsByTagName('type')[0]
#
#     format = movie.


# 进行xml写入

# def get_page(path):
#     """获取给定excel的行数和列数，并返回一个列表[row, col]"""
#     excel = win32com.client.DispatchEx("Excel.Application")
#     excel.Visible = 0
#     excel.DisplayAlerts = True
#     work = excel.Workbooks.Open(path)
#     sheet = work.Worksheets(1)
#     nrows = sheet.UsedRange.Rows.Count
#     ncols = sheet.UsedRange.Columns.Count
#     adr = []
#     if "检表" in work.Name:
#         for row in range(nrows, 0, -1):
#             if check_row(sheet, row, "签名"):
#                 adr.append(int(row)+1)
#                 break
#         for col in range(ncols, 0, -1):
#             if check_col(sheet, col, "检表"):
#                 adr.append(int(col)+2)
#                 break
#     elif "外观记录表" in work.Name:
#         for row in range(nrows, 0, -1):
#             if check_row(sheet, row, "专业监理工程师："):
#                 adr.append(int(row)+1)
#                 break
#         for col in range(ncols, 0, -1):
#             if check_col(sheet, col, "外观记录表"):
#                 adr.append(int(col)+2)
#                 break
#     elif ("记录表" in work.Name) and ("鉴定" not in work.Name) and ("封面" not in work.Name):
#         if "(SG)"in work.Name:
#             for row in range(nrows, 0, -1):
#                 if check_row(sheet, row, "检查人："):
#                     adr.append(int(row)+1)
#                     break
#             for col in range(ncols, 0, -1):
#                 if check_col(sheet, col, "外观记录表"):
#                     adr.append(int(col)+2)
#                     break

def get_page(work, path):
    """获取指定excel文件的行数和列数，并返回一个列表"""
    pos = []
    if os.path.splitext(path)[1] in [".xlsx", ".xls"]:
        # excel = win32com.client.DispatchEx("Excel.Application")
        # excel.Visible = False
        # excel.DisplayAlerts = False
        # work = excel.Workbooks.Open(path)
        sheet = work.Worksheets(1)

        nrows = sheet.UsedRange.Rows.Count
        pos.append(nrows)
        ncols = sheet.UsedRange.Columns.Count
        pos.append(ncols)

        # work.Close()
        # excel.Quit()
    else:
        print("{}无法用excel打开".format(os.path.basename(path)))
    return pos


def read_excel_keys(work, path, list_name):
    """读取文件中的keys值，返回与list_name相对应的列表"""
    # excel = win32com.client.DispatchEx("Excel.Application")
    # excel.Visible = 0
    # excel.DisplayAlerts = 0
    # work = excel.Workbooks.Open(path)
    sheet = work.Worksheets(1)
    nrows = sheet.UsedRange.Rows.Count
    ncols = sheet.UsedRange.Columns.Count
    keys_list = []
    print(work.Name)

    if list_name == "引用":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[0:3] == "%QC":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在引用单元格")
    elif list_name == "日期":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%DT":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在日期单元格")
    elif list_name == "当前页":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%CP":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在当前页单元格")
    elif list_name == "所有页":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%AP":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在所有页单元格")
    elif list_name == "文本框":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%TC":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在文本框单元格")
    elif list_name == "复选框":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%CH":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在复选框单元格")
    elif list_name == "下拉框":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%CB":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在下拉框单元格")
    elif list_name == "合格率":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%EP":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在带合格率的测点单元格")
    elif list_name == "测点":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%MP":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在测点单元格")
    elif list_name == "数字":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%NM":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在数字单元格")
    elif list_name == "时间":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%TI":
                    keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在时间单元格")
    elif list_name == "图片":
        for row in range(1, nrows+1):
            for col in range(1, ncols+1):
                if str(sheet.Cells(row, col).Value).strip()[:3] == "%PT":
                    for num in range(row-1, 0, -1):
                        if ("图" in str(sheet.Cells(num, col).Value)) or \
                                ("照片" in str(sheet.Cells(num, col).Value)):
                                keys_list.append(sheet.Cells(row, col).Value)
        if not keys_list:
            print("不存在图片单元格")

    else:
        print("查询内容不存在：", list_name)
    # work.Close()
    # excel.Quit()

    return keys_list


def read_excel_sign(work, path):
    """提取出所给excel中的签名keys值，返回特定的字典"""
    sign_keys = {}
    # time.sleep(3)
    # excel = win32com.client.DispatchEx("Excel.Application")
    # excel.Visible = 0
    # excel.DisplayAlerts = 0
    # work = excel.Workbooks.Open(path)
    sheet = work.Worksheets(1)
    nrows = sheet.UsedRange.Rows.Count
    ncols = sheet.UsedRange.Columns.Count
    keys_list = []
    print(work.Name)
    try:
        if "外观记录表" in work.Name:
            for row in range(1, nrows+1):
                for col in range(1, ncols+1):
                    if str(sheet.Cells(row, col).Value).strip()[:3] == "%PT":
                        for num in range(col-1, 0, -1):
                            if str(sheet.Cells(row, num).Value).strip() in \
                                    ["签名：", "检查人：", "质检负责人：", "监理员：", "专业监理工程师签名："
                                     "质检人：", "专业监理工程师：", "记录：", "复核：", "总监理工程师：", "质检负责人签名："
                                     ] or ((sheet.Cells(row, num).Value is not None) and
                                           (sheet.Cells(row, num).Value[-3:] in ["签名："])):
                                sign_keys[str(sheet.Cells(row, num).Value)] = str(sheet.Cells(row, col).Value)
                                keys_list.append(sheet.Cells(row, col).Value)
                                break
        elif "检表" in work.Name:
            if "(SG)" in work.Name:
                sg_adr = find_adr(path, "质检负责人签名：")
                jl_adr = find_adr(path, "专业监理工程师签名：")
                for col in range(1, ncols+1):
                    if sg_adr:
                        if str(sheet.Cells(sg_adr[0], col).Value).strip()[:3] == "%PT":
                            for num in range(col - 1, 0, -1):
                                if str(sheet.Cells(sg_adr[0], num).Value).strip() in \
                                        ["签名：", "检查人：", "质检负责人：", "监理员：", "专业监理工程师签名："
                                                                          "质检人：", "专业监理工程师：", "记录：", "复核：", "总监理工程师：", "质检负责人签名："
                                         ] or ((sheet.Cells(sg_adr[0], num).Value is not None) and
                                               (sheet.Cells(sg_adr[0], num).Value[-3:] in ["签名："])):
                                        sign_keys["施工{}".format(str(sheet.Cells(sg_adr[0], num).Value))] = str(sheet.Cells(sg_adr[0], col).Value)
                                        break
                    if jl_adr:
                        if str(sheet.Cells(jl_adr[0], col).Value).strip()[:3] == "%PT":
                            for num in range(col - 1, 0, -1):
                                if str(sheet.Cells(jl_adr[0], num).Value).strip() in \
                                        ["签名：", "检查人：", "质检负责人：", "监理员：", "专业监理工程师签名："
                                                                          "质检人：", "专业监理工程师：", "记录：", "复核：", "总监理工程师：", "质检负责人签名："
                                         ] or ((sheet.Cells(jl_adr[0], num).Value is not None) and
                                               (sheet.Cells(jl_adr[0], num).Value[-3:] in ["签名："])):
                                        sign_keys["监理{}".format(str(sheet.Cells(jl_adr[0], num).Value))] = str(sheet.Cells(jl_adr[0], col).Value)
                                        break
            elif "(JL)" in work.Name:
                cel_1 = sheet.UsedRange.Find("签名：")
                adr_1, adr_2 = [], []
                if cel_1:
                    adr_1 = adr_to_pos(cel_1.Address)
                cel_2 = sheet.UsedRange.FindNext(cel_1)
                if cel_2:
                    adr_2 = adr_to_pos(cel_2.Address)
                if adr_1 and adr_2:
                    for col in range(1, ncols+1):
                        if str(sheet.Cells(adr_1[0], col).Value).strip()[:3] == "%PT":
                            for num in range(col - 1, 0, -1):
                                if str(sheet.Cells(adr_1[0], num).Value).strip() in \
                                        ["签名：", "检查人：", "质检负责人：", "监理员：", "专业监理工程师签名："
                                                                          "质检人：", "专业监理工程师：", "记录：", "复核：", "总监理工程师：", "质检负责人签名："
                                         ] or ((sheet.Cells(adr_1[0], num).Value is not None) and
                                               (sheet.Cells(adr_1[0], num).Value[-3:] in ["签名："])):
                                        sign_keys["监理1{}".format(str(sheet.Cells(adr_1[0], num).Value))] = str(sheet.Cells(adr_1[0], col).Value)
                                        break
                        if str(sheet.Cells(adr_2[0], col).Value).strip()[:3] == "%PT":
                            for num in range(col - 1, 0, -1):
                                if str(sheet.Cells(adr_2[0], num).Value).strip() in \
                                        ["签名：", "检查人：", "质检负责人：", "监理员：", "专业监理工程师签名："
                                                                          "质检人：", "专业监理工程师：", "记录：", "复核：", "总监理工程师：", "质检负责人签名："
                                         ] or ((sheet.Cells(adr_2[0], num).Value is not None) and
                                               (sheet.Cells(adr_2[0], num).Value[-3:] in ["签名："])):
                                        sign_keys["监理2{}".format(str(sheet.Cells(adr_2[0], num).Value))] = str(sheet.Cells(adr_2[0], col).Value)
                                        break
        elif ("记录表" in work.Name) and ("鉴定" not in work.Name) and ("封面" not in work.Name):
            for row in range(1, nrows+1):
                for col in range(1, ncols+1):
                    if str(sheet.Cells(row, col).Value).strip()[:3] == "%PT":
                        for num in range(col-1, 0, -1):
                            if str(sheet.Cells(row, num).Value).strip() in \
                                    ["签名：", "检查人：", "质检负责人：", "监理员：", "专业监理工程师签名："
                                     "质检人：", "专业监理工程师：", "记录：", "复核：", "总监理工程师：", "质检负责人签名："
                                     ] or ((sheet.Cells(row, num).Value is not None) and
                                           (sheet.Cells(row, num).Value[-3:] in ["签名："])):
                                sign_keys[str(sheet.Cells(row, num).Value)] = str(sheet.Cells(row, col).Value)
                                keys_list.append(sheet.Cells(row, col).Value)
                                break

        # work.Close()
        # excel.Quit()
        if not sign_keys:
            print("不存在签名单元格")
            return sign_keys
        else:
            return sign_keys
    except:
        # work.Close()
        # excel.Quit()
        raise Exception


def add_item(dom, par_node, unit, addtions):
    """添加子元素并标明注释,并返回item"""
    base_info = dom.createComment(addtions)
    comment = []
    for child in dom.getElementsByTagName("items")[0].childNodes:
        if child.nodeType == dom.COMMENT_NODE:
            comment.append(child.data)
    if addtions not in "".join(comment):
        par_node.appendChild(base_info)
    item = dom.createElement("item")
    item.setAttribute("Id", str(unit))
    par_node.appendChild(item)
    return item


def excel_write_xml(work, path):
    """根据excel中的keys值生成xml文件,并保存在该路径下"""
    keys = []
    quote_keys = read_excel_keys(work, path, "引用")
    keys.append(quote_keys)
    cur_page_keys = read_excel_keys(work, path, "当前页")
    keys.append(cur_page_keys)
    all_page_keys = read_excel_keys(work, path, "所有页")
    keys.append(all_page_keys)
    date_keys = read_excel_keys(work, path, "日期")
    keys.append(date_keys)
    evaluate_keys = read_excel_keys(work, path, "合格率")
    keys.append(evaluate_keys)
    measure_keys = read_excel_keys(work, path, "测点")
    keys.append(measure_keys)
    txt_keys = read_excel_keys(work, path, "文本框")
    keys.append(txt_keys)
    number_keys = read_excel_keys(work, path, "数字")
    keys.append(number_keys)
    time_keys = read_excel_keys(work, path, "时间")
    keys.append(time_keys)
    check_keys = read_excel_keys(work, path, "复选框")
    keys.append(check_keys)
    com_box_keys = read_excel_keys(work, path, "下拉框")
    keys.append(com_box_keys)
    pic_keys = read_excel_keys(work, path, "图片")
    keys.append(pic_keys)

    xmldom = Document()
    pos = get_page(work, path)
    grid = xmldom.createElement("grid")
    items = xmldom.createElement("items")
    scope = xmldom.createElement("range")
    ran_des = xmldom.createComment("表单范围")
    tab_name = xmldom.createComment(os.path.basename(path))
    # row = dom.createAttribute("maxRows")
    # row.nodeValue = pos[0]
    # scope.setAttributeNode(row)
    xmldom.appendChild(grid)
    grid.appendChild(items)
    grid.appendChild(ran_des)
    scope.setAttribute("maxRows", str(pos[0]+1))
    scope.setAttribute("maxCols", str(pos[1]+1))
    if ("记录表" in os.path.basename(path)) and ("鉴定" not in os.path.basename(path)):
        sheet = xmldom.createElement("sheet")
        sheet.setAttribute("Authority", "1111")
        grid.appendChild(sheet)
    grid.appendChild(scope)
    items.appendChild(tab_name)

    for key in keys:
        if key:
            for unit in key:
                if str(unit)[:3] == "%QC":
                    item = add_item(xmldom, items, unit, "表单基础信息")
                    item.setAttribute("Type", "QuoteCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%CP":
                    item = add_item(xmldom, items, unit, "当前页")
                    item.setAttribute("Type", "CurPageCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%AP":
                    item = add_item(xmldom, items, unit, "所有页")
                    item.setAttribute("Type", "AllPageCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%DT":
                    item = add_item(xmldom, items, unit, "表单日期")
                    item.setAttribute("Type", "DateCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%EP":
                    item = add_item(xmldom, items, unit, "带合格率的测点单元格")
                    item.setAttribute("Type", "EvaluateCell")
                    item.setAttribute("Enable", "true")
                    item.setAttribute("ShowType", "1")
                    item.setAttribute("Text", "应测 %d 点（处），实测 %d 点（处），合格 %d 点（处），合格率 %.1f %%，数据详见检表")
                elif str(unit)[:3] == "%MP":
                    item = add_item(xmldom, items, unit, "测点单元格")
                    item.setAttribute("Type", "MeasureCell")
                    item.setAttribute("InputType", "ArrayNumber")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%TC":
                    item = add_item(xmldom, items, unit, "文本内容")
                    if ("记录表" in os.path.basename(path)) and ("外观" not in os.path.basename(path)):
                        write_formula(work, path, unit, item)
                    item.setAttribute("Type", "TextCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%NM":
                    item = add_item(xmldom, items, unit, "数字单元格")
                    item.setAttribute("Type", "NumberCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%TI":
                    item = add_item(xmldom, items, unit, "时间单元格")
                    item.setAttribute("Type", "TimeCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%CH":
                    item = add_item(xmldom, items, unit, "表单复选框")
                    item.setAttribute("Type", "CheckBoxCell")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%CB":
                    item = add_item(xmldom, items, unit, "表单下拉框")
                    item.setAttribute("Type", "ComboBoxCell")
                    item.setAttribute("Options", " @@合格@@不合格")
                    item.setAttribute("Enable", "true")
                elif str(unit)[:3] == "%PT":
                    item = add_item(xmldom, items, unit, "表单图片")
                    item.setAttribute("Type", "PictureCell")
                    item.setAttribute("Enable", "true")
    sign_keys = read_excel_sign(work, path)
    if sign_keys:
        keys_lists = list(sign_keys.keys())
        sign_com = xmldom.createComment("表单签名签章")
        sign_info = xmldom.createComment("签名权限码40001001业主 40001002监理 40001003施工 Role自动签名角色名称")
        items.appendChild(sign_com)
        items.appendChild(sign_info)

        for keys in keys_lists:
            if ("施工" in keys) or (keys in ["质检人：", "检查人：", "质检负责人：", "质检人签名：", "检查人签名：",
                                           "质检负责人签名：", "检验负责人：", "检测：", "记录：", "复核："]):
                item = xmldom.createElement("item")
                item.setAttribute("Sign", "40001003")
                item.setAttribute("Id", sign_keys[keys])
                item.setAttribute("Type", "PictureCell")
                item.setAttribute("Enable", "true")
                if (("施工" in keys) and ("质检负责人" not in keys)) or (keys in ["质检人：", "检查人：",  "质检人签名：", "检查人签名：",
                                                                           "检测：", "记录：", "复核："]):
                    item.setAttribute("Role", "检查人")
                elif "质检负责人" in keys:
                    item.setAttribute("Role", "质检负责人")
                items.appendChild(item)
            elif ("监理" in keys) or (keys in ["监理员：", "监理员签名：", "专业监理工程师：", "专业监理工程师签名：",
                                             "总监理工程师：", "总监理工程师签名："]):
                item = xmldom.createElement("item")
                item.setAttribute("Sign", "40001002")
                item.setAttribute("Id", sign_keys[keys])
                item.setAttribute("Type", "PictureCell")
                item.setAttribute("Enable", "true")
                if ("监理" in keys) and ("监理1" not in keys) and ("监理2" not in keys) and ("专业监理工程师" not in keys):
                    item.setAttribute("Role", "监理员")
                elif "监理1" in keys:
                    item.setAttribute("Role", "监理员")
                elif "监理2" in keys:
                    item.setAttribute("Role", "专业监理工程师")
                else:
                    item.setAttribute("Role", keys)
                items.appendChild(item)

    f = open("{}.xml".format(os.path.splitext(path)[0]), encoding='utf-8', mode='w')
    xmldom.writexml(f, addindent="\t", newl="\n", encoding="UTF-8")
    # f.write(str(xmldom.toxml(encoding="utf-8")))
    f.close()
    print("{}生成xml文件完成。".format(os.path.basename(path)))


def write_formula(work, path, unit, item):
    """在xml中写入公式"""

    sheet = work.Worksheets(1)
    status = False
    first = ""
    second = ""

    adr = find_adr(path, unit)
    for row in range(int(adr[0])-1, 0, -1):
        if sheet.Cells(row, adr[1]).Value in ["差值（cm）", "差值（mm）", "差值{}（mm）".format("\n"), "偏差（mm）"]:
            status = True
            break
    if status:
        for col in range(int(adr[1])-1, 0, -1):
            if first and second:
                break
            else:
                if str(sheet.Cells(adr[0], col).Value).strip()[:3] == "%TC":
                    if not first:
                        for row in range(int(adr[0]) - 1, 0, -1):
                            if "实测" in str(sheet.Cells(row, col).Value):
                                first = sheet.Cells(adr[0], col).Value
                                break
                    if not second:
                        for row in range(int(adr[0]) - 1, 0, -1):
                            if "设计" in str(sheet.Cells(row, col).Value):
                                second = sheet.Cells(adr[0], col).Value
                                break
    if first and second:
        for_str = first+"-"+second
        # print(for_str)
        item.setAttribute("Formula", for_str)


path_lists = [r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\2 路面检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\2 路面检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\2 路面检表\施工检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\2 路面检表\施工检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\3 路面外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\3 路面外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\3 路面外观记录表\施工外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\3 路面外观记录表\施工外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\4 路面记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表\施工记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\4 路面记录表\施工记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\6 交通安全设施（第六册）\2 交安工程监理单位用表\2 交安检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\6 交通安全设施（第六册）\1 交安工程施工单位用表\2 交安检表\检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\6 交通安全设施（第六册）\2 交安工程监理单位用表\2 交安检表\检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\7 绿化工程（第七册）\2 绿化工程用表（监理单位使用）\2 绿化检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\7 绿化工程（第七册）\1 绿化工程用表（施工单位使用）\2 绿化检表\检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\7 绿化工程（第七册）\2 绿化工程用表（监理单位使用）\2 绿化检表\检表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\8 声屏障工程（第八册）\2 声屏障工程用表（监理单位使用）\3 声屏障外观鉴定记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\8 声屏障工程（第八册）\1 声屏障工程用表（施工单位使用）\3 声屏障外观鉴定记录表\外观记录表",
              r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\8 声屏障工程（第八册）\2 声屏障工程用表（监理单位使用）\3 声屏障外观鉴定记录表\外观记录表",
              ]

paths = [r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\1 路面工程施工单位用表\4 路面记录表\施工记录表",
         r"C:\Users\admin\Desktop\替换筑业\第三批\二稿\3 路面工程（第三册）\3 路面工程\2 路面工程监理单位用表\4 路面记录表\施工记录表",
         ]

# if __name__ == "__main__":
#     for unit in paths:
#         paths = os.listdir(unit)
#         for item in paths:
#             path = os.path.join(unit, item)
#             if os.path.splitext(path)[1] == ".xlsx":
#                 time.sleep(3)
#                 excel = win32com.client.DispatchEx("Excel.Application")
#                 excel.Visible = 0
#                 excel.DisplayAlerts = 0
#                 work = excel.Workbooks.Open(path)
#
#                 excel_write_xml(work, path)
#
#                 work.Close()
#                 excel.Quit()
path = r"C:\Users\admin\Desktop\03 路面厚度现场质量检查记录表（总厚度）(SG).xlsx"
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = 0
excel.DisplayAlerts = 0
work = excel.Workbooks.Open(path)

# sheet = work.Worksheets(1)
# status = False
# first = ""
# second = ""
#
# adr = find_adr(path, "%TCJ09")
# print(adr)
# for row in range(int(adr[0])-1, 0, -1):
#     print(row)
#     if sheet.Cells(row, adr[1]).Value in ["差值（cm）", "差值（mm）", "差值{}（mm）".format("\n"), "偏差（mm）"]:
#         status = True
#         print(True)
#         break
# if status:
#     for col in range(int(adr[1])-1, 0, -1):
#         if first and second:
#             break
#         else:
#             if str(sheet.Cells(adr[0], col).Value).strip()[:3] == "%TC":
#                 if not first:
#                     for row in range(int(adr[0])-1, 0, -1):
#                         if "实测" in str(sheet.Cells(row, col).Value):
#                             first = sheet.Cells(adr[0], col).Value
#                             break
#                 if not second:
#                     for row in range(int(adr[0]) - 1, 0, -1):
#                         if "设计" in str(sheet.Cells(row, col).Value):
#                             second = sheet.Cells(adr[0], col).Value
#                             break
#
# if first and second:
#     for_str = first+"-"+second
#     print(for_str)
excel_write_xml(work, path)
work.Close()
excel.Quit()
