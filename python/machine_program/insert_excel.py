#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-03-21 09:40
# Author  : MrFiona
# File    : insert_excel.py
# Software: PyCharm Community Edition

import sys
reload(sys)
import time
import xlwt
sys.setdefaultencoding('utf-8')
import xlrd
from xlutils.copy import copy
import xlwt
# import get_word
import xlsxwriter

# url = 'https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/Silver/2017%20WW11/5691_Silver.html'
# data_type = 'HW Configuration'
# #插入excel表格中的HW工作表
# num = 1
# start = time.time()
#
#
# def return_label_num_list(label_num_list=None):
#     if not label_num_list:
#         label_num_list = []
#     # 写入多个周的数据
#     for i in range(50):
#         # i为偶数
#         if i & 1 == 0:
#             if i < 2:
#                 x = ['D', 'Q']
#             elif i >= 2:
#                 prefix = chr(ord('A') + (i - 2) / 2)
#                 x = [prefix + 'D', prefix + 'Q']
#         elif i & 1 == 1:
#             if i < 3:
#                 x = ['Q', 'AD']
#             elif i > 3:
#                 prefix_font = chr(ord('A') + (i - 3) / 2)
#                 prefix_behind = chr(ord('A') + (i - 3) / 2 + 1)
#                 x = [prefix_font + 'Q', prefix_behind + 'D']
#             elif i == 3:
#                 x = ['AQ', 'BD']
#         # if i & 1 == 0:
#         print 'i = %d, x = %s' % (i, x)
#
# #创建一个新的excel工作簿  文件对象
# workbook = xlsxwriter.Workbook('report_result.xlsx')
# #在excel工作簿中增加工作表单
# worksheet = workbook.add_worksheet('HW')
#
# label_num_list = return_label_num_list()
#
# for j in range(2):
#     #为excel对象设定显示格式
#     bold = workbook.add_format({'bold': 1, 'align':'center', 'font_size': 12})
#     bold_merge = workbook.add_format({'bold': 2, 'align':'justify', 'valign':'vcenter' ,'bg_color': '#C6EFCE'})
#     #获取提取的待插入excel表的数据
#     obj = get_word.GetAnalysisData(url, data_type)
#     row_coordinates_list, column_coordinates_list, cell_data_list, header_cell_data_info = obj.get_hw_data()
#     # print row_coordinates_list
#     # print column_coordinates_list
#     # print cell_data_list, len(cell_data_list)
#     # print header_cell_data_info, len(header_cell_data_info)
#     #设置第一行单元格的宽度为20，字体为bold格式
#     # worksheet.set_row(0, 40, bold)
#     # for num in range(1,10):
#     #     worksheet.set_row(num, 40)
#     #写入工作表的第一行,现在只支持system 1 到system 10，截取10行数据
#     #为当前workbook添加一个样式名为titleformat
#     titleformat = workbook.add_format()
#     titleformat_header = workbook.add_format()
#     titleformat_header.set_bold() # 设置粗体字
#     titleformat.set_font_size(12) # 设置字体大小为10
#     titleformat_header.set_font_size(12) # 设置字体大小为10
#     titleformat.set_font_name('Microsoft yahei') # 设置字体样式为雅黑
#     titleformat.set_align('center') # 设置水平居中对齐
#     titleformat_header.set_align('center') # 设置水平居中对齐
#     titleformat.set_align('vcenter')  # 设置垂直居中对齐
#     titleformat_header.set_align('vcenter')  # 设置垂直居中对齐
#     # def fun(worksheet, label):
#     #     def convert_chr(asc,num):
#     #         return chr( ord(asc) + num )
#
#     worksheet.write('{0}1' .format( chr(ord('A')+13*j) ), 'HW')
#     worksheet.write_row('{0}1' .format( chr(ord('A')+ 3 + 13*j) ), row_coordinates_list[:10])
#     #写入工作表的第二行
#     if j == 0:
#         common_formula_lable_list = [['D', 'Q'], ['E', 'R'], ['F', 'S'], ['G', 'T'], ['H', 'U'],
#                                  ['I', 'V'], ['J', 'W'], ['K', 'X'], ['L', 'Y'], ['M', 'Z']]
#     if j == 1:
#         common_formula_lable_list = [['Q', 'AD'], ['R', 'AE'], ['S', 'AF'], ['T', 'AG'], ['U', 'AH'],
#                                  ['V', 'AI'], ['W', 'AJ'], ['X', 'AL'], ['Y', 'AM'], ['Z', 'AN']]
#     insert_row_formula_list = []
#     for i in range(len(common_formula_lable_list)):
#         if j == 0:
#             row_formula = '=OR(IF(OR(ISBLANK({0}3),ISBLANK({1}3)),FALSE,NOT(EXACT({0}3,{1}3))),IF(OR(ISBLANK({0}4),ISBLANK({1}4)),FALSE,NOT(EXACT({0}4,{1}4))),IF(OR(ISBLANK({0}5),ISBLANK({1}5)),FALSE,NOT(EXACT({0}5,{1}5))),IF(OR(ISBLANK({0}6),ISBLANK({1}6)),FALSE,NOT(EXACT({0}6,{1}6))),IF(OR(ISBLANK({0}7),ISBLANK({1}7)),FALSE,NOT(EXACT({0}7,{1}7))),IF(OR(ISBLANK({0}8),ISBLANK({1}8)),FALSE,NOT(EXACT({0}8,{1}8))),IF(OR(ISBLANK({0}9),ISBLANK({1}9)),FALSE,NOT(EXACT({0}9,{1}9))))'.format(common_formula_lable_list[i][0], common_formula_lable_list[i][1])
#         if j == 1:
#             row_formula = '=OR(IF(OR(ISBLANK(Q3),ISBLANK({1}3)),FALSE,NOT(EXACT({0}3,{1}3))),IF(OR(ISBLANK({0}4),ISBLANK({1}4)),FALSE,NOT(EXACT({0}4,{1}4))),IF(OR(ISBLANK({0}5),ISBLANK({1}5)),FALSE,NOT(EXACT({0}5,{1}5))),IF(OR(ISBLANK({0}6),ISBLANK({1}6)),FALSE,NOT(EXACT({0}6,{1}6))),IF(OR(ISBLANK({0}7),ISBLANK({1}7)),FALSE,NOT(EXACT({0}7,{1}7))),IF(OR(ISBLANK({0}8),ISBLANK({1}8)),FALSE,NOT(EXACT({0}8,{1}8))),IF(OR(ISBLANK({0}9),ISBLANK({1}9)),FALSE,NOT(EXACT({0}9,{1}9))))'.format(common_formula_lable_list[i][0], common_formula_lable_list[i][1])
#
#         # print row_formula
#         insert_row_formula_list.append(row_formula)
#     worksheet.write_row('{0}2' .format( chr(ord('A')+ 2 + 13*j) ), ['Changed from previous configuration?'] + insert_row_formula_list)
#     #按列插入数据
#     insert_col_formula_list = []
#     worksheet.write_column('{0}3' .format( chr(ord('A')+ 1 + 13*j) ), column_coordinates_list[2:])
#     #设置D-M列的每一个单元格的宽度为60
#     # worksheet.set_column('D:M', 65)
#     # worksheet.set_column('C:C', 40)
#     # worksheet.set_column('B:B', 20)
#     for col in range(3, 10):
#         if j == 0:
#             column_formula = '=OR(IF(OR(ISBLANK(D{0}),ISBLANK(Q{0})),FALSE,NOT(EXACT(D{0},Q{0}))),IF(OR(ISBLANK(E{0}),ISBLANK(R{0})),FALSE,NOT(EXACT(E{0},R{0}))),IF(OR(ISBLANK(F{0}),ISBLANK(S{0})),FALSE,NOT(EXACT(F{0},S{0}))),IF(OR(ISBLANK(G{0}),ISBLANK(T{0})),FALSE,NOT(EXACT(G{0},T{0}))),IF(OR(ISBLANK(H{0}),ISBLANK(U{0})),FALSE,NOT(EXACT(G{0},T{0}))),IF(OR(ISBLANK(H{0}),ISBLANK(U{0})),FALSE,NOT(EXACT(H{0},U{0}))),IF(OR(ISBLANK(I{0}),ISBLANK(V{0})),FALSE,NOT(EXACT(I{0},V{0}))),IF(OR(ISBLANK(J{0}),ISBLANK(W{0})),FALSE,NOT(EXACT(J{0},W{0}))),IF(OR(ISBLANK(K{0}),ISBLANK(X{0})),FALSE,NOT(EXACT(K{0},X{0}))))' .format(col)
#         if j == 1:
#             column_formula = '=OR(IF(OR(ISBLANK(Q{0}),ISBLANK(AD{0})),FALSE,NOT(EXACT(Q{0},AD{0}))),IF(OR(ISBLANK(R{0}),ISBLANK(AE{0})),FALSE,NOT(EXACT(R{0},AE{0}))),IF(OR(ISBLANK(S{0}),ISBLANK(AF{0})),FALSE,NOT(EXACT(S{0},AF{0}))),IF(OR(ISBLANK(T{0}),ISBLANK(AG{0})),FALSE,NOT(EXACT(T{0},AG{0}))),IF(OR(ISBLANK(U{0}),ISBLANK(AH{0})),FALSE,NOT(EXACT(T{0},AG{0}))),IF(OR(ISBLANK(U{0}),ISBLANK(AH{0})),FALSE,NOT(EXACT(U{0},AH{0}))),IF(OR(ISBLANK(V{0}),ISBLANK(AI{0})),FALSE,NOT(EXACT(V{0},AI{0}))),IF(OR(ISBLANK(W{0}),ISBLANK(AJ{0})),FALSE,NOT(EXACT(W{0},AJ{0}))),IF(OR(ISBLANK(X{0}),ISBLANK(AK{0})),FALSE,NOT(EXACT(X{0},AK{0}))))' .format(col)
#         # print column_formula
#         insert_col_formula_list.append(column_formula)
#     worksheet.write_column('{0}3' .format( chr(ord('A')+ 2 + 13*j) ), insert_col_formula_list)
#     #按行插入数据
#     for row_num in range(3,10):
#         #每个元素限制为10列 [:10]
#         worksheet.write_row('{1}{0}'.format(row_num, chr(ord('A')+ 3 + 13*j)), cell_data_list[row_num-1][:10])
#     #合并单元格
#     worksheet.merge_range('{0}3:{0}9' .format( chr(ord('A')+ 13*j) ), "")
#     # Set up some formats to use.
#     red = workbook.add_format({'color': 'red'})
#     worksheet.write_rich_string('{0}3' .format( chr(ord('A')+ 13*j) ), red, 'WW11', titleformat)
#
#     first_column_label_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
#     num -= 1
#
# workbook.close()
# print time.time() - start
#
# #读取excel表格
# data = xlrd.open_workbook('report_result11.xls', formatting_info=True)
# table = data.sheet_by_name('HW')
# nrows = table.nrows
# ncols = table.ncols
# print nrows, ncols
#
# wtable = copy(data)
# wtable.add_sheet('gggg')
# wtable.sheet_index('HW')
#
# wtable.save('hhhw.xlsx')
#
#
#
# import xlwt
# workbook = xlwt.Workbook()
# worksheet = workbook.add_sheet('My Sheet')
# pattern = xlwt.Pattern() # Create the Pattern
# pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
# pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
# style = xlwt.XFStyle() # Create the Pattern
# style.pattern = pattern # Add Pattern to Style
# worksheet.write(0, 0, 'Cell Contents', style)
# workbook.save('Excel_Workbook.xls')

import xlrd
import xlutils.copy

newwb = xlrd.open_workbook('dd.xls', formatting_info=True)  # formatting_info 带格式导入
outwb = xlutils.copy.copy(newwb)  # 建立一个副本来用xlwt来写

# 修改值

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """

    def _getOutCell(outSheet, colIndex, rowIndex):
        """ HACK: Extract the internal xlwt cell representation. """
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row: return None

        cell = row._Row__cells.get(colIndex)
        return cell

    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
            # END HACK


outSheet = outwb.get_sheet(0)
setOutCell(outSheet, 5, 5, 'Test')
outwb.save('output.xls')




import xlrd, xlwt

# Demonstration of copy2 patch for xlutils 1.4.1

# Context:
# xlutils.copy.copy(xlrd_workbook) -> xlwt_workbook
# copy2(xlrd_workbook) -> (xlwt_workbook, style_list)
# style_list is a conversion of xlrd_workbook.xf_list to xlwt-compatible styles

# Step 1: Create an input file for the demo
def create_input_file():
    wtbook = xlwt.Workbook()
    wtsheet = wtbook.add_sheet(u'First')
    colours = 'white black red green blue pink turquoise yellow'.split()
    fancy_styles = [xlwt.easyxf(
        'font: name Times New Roman, italic on;'
        'pattern: pattern solid, fore_colour %s;'
         % colour) for colour in colours]
    for rowx in xrange(8):
        wtsheet.write(rowx, 0, rowx)
        wtsheet.write(rowx, 1, colours[rowx], fancy_styles[rowx])
    wtbook.save('demo_copy2_in.xls')

# Step 2: Copy the file, changing data content
# ('pink' -> 'MAGENTA', 'turquoise' -> 'CYAN')
# without changing the formatting

from xlutils.filter import process,XLRDReader,XLWTWriter

# Patch: add this function to the end of xlutils/copy.py
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list

def update_content():
    rdbook = xlrd.open_workbook('22222.xls', formatting_info=True)
    sheetx = 0
    rdsheet = rdbook.sheet_by_index(sheetx)
    wtbook, style_list = copy2(rdbook)
    wtsheet = wtbook.get_sheet(sheetx)
    fixups = [(5, 1, 'MAGENTA'), (6, 1, 'CYAN')]
    for rowx, colx, value in fixups:
        xf_index = rdsheet.cell_xf_index(rowx, colx)
        wtsheet.write(rowx, colx, value, style_list[xf_index])
    wtbook.save('demo_copy2_out.xls')

# create_input_file()
update_content()

# workbook = xlwt.Workbook()
# worksheet = workbook.add_sheet('ee')
# # print 'ww', xlrd.empty_cell.value
# insert_col_formula_list = []
# worksheet.write(1,2,'=OR(IF(OR(ISBLANK(D{0}),ISBLANK(Q{0})),FALSE,NOT(EXACT(D{0},Q{0}))),IF(OR(ISBLANK(E{0}),ISBLANK(R{0})),FALSE,NOT(EXACT(E{0},R{0}))),IF(OR(ISBLANK(F{0}),ISBLANK(S{0})),FALSE,NOT(EXACT(F{0},S{0}))),IF(OR(ISBLANK(G{0}),ISBLANK(T{0})),FALSE,NOT(EXACT(G{0},T{0}))),IF(OR(ISBLANK(H{0}),ISBLANK(U{0})),FALSE,NOT(EXACT(G{0},T{0}))),IF(OR(ISBLANK(H{0}),ISBLANK(U{0})),FALSE,NOT(EXACT(H{0},U{0}))),IF(OR(ISBLANK(I{0}),ISBLANK(V{0})),FALSE,NOT(EXACT(I{0},V{0}))),IF(OR(ISBLANK(J{0}),ISBLANK(W{0})),FALSE,NOT(EXACT(J{0},W{0}))),IF(OR(ISBLANK(K{0}),ISBLANK(X{0})),FALSE,NOT(EXACT(K{0},X{0}))))' .format(3))
# workbook.save('dd.xls')

#A2, B2, C2, D2, E2, F2, G2, H2, I2, J2, K2, L2, M2
#N2, O2, P2, Q2, R2, S2, T2, U2, V2, W2, X2, Y2, Z2
#AA2, AB2, AC2, AD2, AE2, AF2, AG2, AH2, AI2, AJ2, AK2, AL2, AM2
#AN2, AO2, AP2, AQ2, AR2, AS2, AT2, AU2, AV2, AW2, AX2, AY2, AZ2


# -*- coding: cp936 -*-
import xlrd
import xlwt

xls1 = xlrd.open_workbook("22222.xls")
mysheet1 = xls1.sheet_by_name("HW")  # 找到名为Sheet1的工作表。区分大小写
print("表1共有 %d 行， %d 列。" % (mysheet1.nrows, mysheet1.ncols))

xls2 = xlwt.Workbook()
mysheet2 = xls2.add_sheet('SheetB')  # 建立2号xls，并往sheetB里面写入数据

# 获取表1的表头
titlelist = mysheet1.row_values(0)

# 每行的操作.生成110个分表
title0 = 'xxxxx事项分表'
titlex = 'xxxx流程图（xxxx）'
mysheet2.write(0, 0, title0.decode('utf-8'))
mysheet2.write(0, 8, titlex.decode('utf-8'))
for x in range(0, 8):  # 表1 共有110行数据
    titlelist1 = mysheet1.row_values(x + 1)  # 每一次循环都得到第x+1行数据。
    x = x * 9
    print x
    for row in range(x + 1, x + 9):  # 每9行为一组 所以上一行的x乘以了9。
        if row == x + 1:  # 如果是 每组 的第一行
            for i in range(0, 3):  # 【此处填写第 x+1行(也就是每组的第一行) 的实际列数】
                mysheet2.write(row, i * 2, titlelist[i])  # 偶数列写入表头数据
                mysheet2.write(row, i * 2 + 1, titlelist1[i])  # 奇数列写入表1的第x+1行的数据
        else:
            if row == x + 2:  # 如果是 每组 的第二行
                mysheet2.write(row, 0, titlelist[3])  # 因为第二行只有两个格子，不用循环，只写第0列和第1列
                mysheet2.write(row, 1, titlelist1[3])
            else:
                if row == x + 3:  # 如果是 每组 的第三行
                    mysheet2.write(row, 0, titlelist[4])  # 第三行也只有两个空，不用循环，只写第0列和第1列
                    mysheet2.write(row, 1, titlelist1[4])

                else:
                    if row == x + 4:
                        for i in range(0, 4):
                            mysheet2.write(row, i * 2, titlelist[i + 5])
                            mysheet2.write(row, i * 2 + 1, titlelist1[i + 5])
                    else:
                        if row == x + 5:
                            for i in range(0, 4):
                                mysheet2.write(row, i * 2, titlelist[i + 9])  # 每写一次，列数向后推4。如5,9,13,17,21…
                                mysheet2.write(row, i * 2 + 1, titlelist1[i + 9])  # 因为“for i in range(0,4):”，一次写4个列
                        else:
                            if row == x + 6:
                                for i in range(0, 4):
                                    mysheet2.write(row, i * 2, titlelist[i + 13])
                                    mysheet2.write(row, i * 2 + 1, titlelist1[i + 13])
                            else:
                                if row == x + 7:
                                    for i in range(0, 4):
                                        mysheet2.write(row, i * 2, titlelist[i + 17])
                                        mysheet2.write(row, i * 2 + 1, titlelist1[i + 17])
                                else:
                                    if row == x + 8:
                                        for i in range(0, 4):  # 如果是 每组 的第八行
                                            mysheet2.write(row, i * 2, titlelist[i + 21])  # 每写一次，列数向后推4。如5,9,13,17,21…
                                            mysheet2.write(row, i * 2 + 1, titlelist1[i + 21])

xls2.save('auto.xls')  # 将填充后的结果保存为auto.xls




















