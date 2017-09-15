# -*- coding: utf-8 -*-
import datetime
import sys,os   
import collections
from optparse import OptionParser
import xlsxwriter  
from xlsxwriter.workbook import Workbook  
from xlrd.sheet import Sheet

#python testReport.py ass 1.0 2017-09-15 50s 500 490 0 10

def get_format(wd, option={}):
    return wd.add_format(option)

# 设置居中
def get_format_center(wd,num=1):
    return wd.add_format({'align': 'center','valign': 'vcenter','border':num})
def set_border_(wd, num=1):
    return wd.add_format({}).set_border(num)

# 写数据
def _write_center(worksheet, cl, data, wd):
    return worksheet.write(cl, data, get_format_center(wd))

def init(worksheet,data):

    # 设置列行的宽高
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 20)
    worksheet.set_column("C:C", 20)
    worksheet.set_column("D:D", 20)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 20)

    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    worksheet.set_row(3, 30)
    worksheet.set_row(4, 30)
    worksheet.set_row(5, 30)

    # worksheet.set_row(0, 200)

    define_format_H1 = get_format(workbook, {'bold': True, 'font_size': 18})
    define_format_H2 = get_format(workbook, {'bold': True, 'font_size': 14})
    define_format_H3 = get_format(workbook, {'bold': True, 'font_size': 20})

    define_format_H1.set_border(1)

    define_format_H2.set_border(1)
    define_format_H1.set_align("center")
    define_format_H2.set_align("center")
    define_format_H2.set_bg_color("blue")
    define_format_H2.set_color("#ffffff")
    define_format_H3.set_color("red")
    define_format_H3.set_align("center")
    define_format_H3.set_border(1)


    # Create a new Chart object.

    worksheet.merge_range('A1:E1', '测试报告总概况', define_format_H1)
    worksheet.merge_range('A2:E2', '测试概括', define_format_H2)

    _write_center(worksheet, "A3", '项目名称', workbook)
    _write_center(worksheet, "A4", '项目版本', workbook)
    _write_center(worksheet, "A5", '测试耗时', workbook)
    _write_center(worksheet, "A6", '测试日期', workbook)
    


    _write_center(worksheet, "B3", data['test_name'], workbook)
    _write_center(worksheet, "B4", data['test_version'], workbook)
    _write_center(worksheet, "B5", data['test_time'], workbook)
    _write_center(worksheet, "B6", data['test_data'], workbook)

    _write_center(worksheet, "C3", '接口总数', workbook)
    _write_center(worksheet, "C4", '通过总数', workbook)
    _write_center(worksheet, "C5", '失败总数', workbook)
    _write_center(worksheet, "C6", '跳过总数', workbook)


    _write_center(worksheet, "D3", int(data['test_sum']), workbook)
    _write_center(worksheet, "D4", int(data['test_success']), workbook)
    _write_center(worksheet, "D5", int(data['test_failed']), workbook)
    _write_center(worksheet, "D6", int(data['test_skip']), workbook)


    _write_center(worksheet, "E3", "结论", workbook)

    pass_percent = (int(data['test_success']) / int(data['test_sum']))*100
    if 	pass_percent > 90.0:
    	test_Conclusion = 'Pass'
    else:test_Conclusion = 'Fail'

    worksheet.merge_range('E4:E6', test_Conclusion, define_format_H3)

    pie(workbook, worksheet)

 # 生成饼形图
def pie(workbook, worksheet):
    chart1 = workbook.add_chart({'type': 'pie'})
    chart1.add_series({
	    'name':       '测试统计',
	    'categories':'=测试总况!$C$4:$C$6',
	    'values':    '=测试总况!$D$4:$D$6',
	    'points':[  
	            {'fill':{'color':'#00CD00'}},  
	            {'fill':{'color':'red'}},  
	            {'fill':{'color':'gray'}},  
	            ], 
    })
    chart1.set_title({'name': '测试统计'})
    chart1.set_style(3)
    worksheet.insert_chart('A9', chart1, {'x_offset': 25, 'y_offset': 10})



if __name__ == '__main__':
    workbook = xlsxwriter.Workbook('aototest测试报告.xlsx')
    worksheet = workbook.add_worksheet("测试总况")

    params = sys.argv[1:]
    data = {}
    data1 = {}

    data['test_name'] = params[0]
    data['test_version'] = params[1]
    data['test_time'] = params[2]
    data['test_data'] = params[3]
    
    data['test_sum'] = params[4]
    data['test_success'] = params[5]
    data['test_failed'] = params[6]
    data['test_skip'] = params[7]

    init(worksheet,data)

    workbook.close()