#coding=utf-8  
import xlsxwriter  
from xlsxwriter.workbook import Workbook  
from xlrd.sheet import Sheet  
  
def demo1():  
    import xlsxwriter  
  
  
    # ����excel�ļ�  
    workbook = xlsxwriter.Workbook('demo.xlsx')  
#     ���worksheet,Ҳ����ָ������  
    worksheet = workbook.add_worksheet()  
    worksheet = workbook.add_worksheet('Test')  
      
    #���õ�һ�еĿ��  
    worksheet.set_column('A:A', len('hello ')+1)  
      
    #���һ���Ӵָ�ʽ�������ʹ��  
    bold = workbook.add_format({'bold': True})  
      
    #��A1��Ԫ��д�봿�ı�  
    worksheet.write('A1', 'Hello')  
      
    #��A2��Ԫ��д�����ʽ���ı�  
    worksheet.write('A2', 'World', bold)  
      
    #ָ������д�����֣��±��0��ʼ  
    worksheet.write(2, 0, 123)  
    worksheet.write(3, 0, 123.456)  
      
    #��B5��Ԫ�����ͼƬ  
    worksheet.insert_image('B5', 'python-logo.png')  
      
      
    workbook.close()  
      
      
def charts():  
    workbook = xlsxwriter.Workbook('chart_column.xlsx')  
    worksheet = workbook.add_worksheet()  
    bold = workbook.add_format({'bold': 1})  
       
    # ���Ǹ�����table����  
    headings = ['Number', 'Batch 1', 'Batch 2']  
    data = [  
        [2, 3, 4, 5, 6, 7],  
        [10, 40, 50, 20, 10, 50],  
        [30, 60, 70, 50, 40, 30],  
    ]  
    #д��һ��   
    worksheet.write_row('A1', headings, bold)  
    #д��һ��  
    worksheet.write_column('A2', data[0])  
    worksheet.write_column('B2', data[1])  
    worksheet.write_column('C2', data[2])  
      
      
       
    ############################################  
    #����һ��ͼ��������column  
    chart1 = workbook.add_chart({'type': 'column'})  
       
    # ����series,�����ǰ��worksheet���й�ϵ�ġ�   
#     ָ��ͼ������ݷ�Χ  
    chart1.add_series({  
        'name':       '=Sheet1!$B$1',  
        'categories': '=Sheet1!$A$2:$A$7',  
        'values':     '=Sheet1!$B$2:$B$7',  
    })  
    chart1.add_series({  
        'name':       "=Sheet1!$C$1",  
        'categories': '=Sheet1!$A$2:$A$7',  
        'values':     '=Sheet1!$C$2:$C$7',  
    })  
#    ����series����һ�ַ���     
#     #     [sheetname, first_row, first_col, last_row, last_col]   
#     chart1.add_series({  
#         'name':         ['Sheet1',0,1],  
#         'categories':   ['Sheet1',1,0,6,0],  
#         'values':       ['Sheet1',1,1,6,1],  
#                        })  
#         
#   
#   
#     chart1.add_series({  
#         'name':       ['Sheet1', 0, 2],  
#         'categories': ['Sheet1', 1, 0, 6, 0],  
#         'values':     ['Sheet1', 1, 2, 6, 2],  
#     })  
      
    
#      ���ͼ�����ͱ�ǩ  
    chart1.set_title ({'name': 'Results of sample analysis'})  
    chart1.set_x_axis({'name': 'Test number'})  
    chart1.set_y_axis({'name': 'Sample length (mm)'})  
       
    # ����ͼ����  
    chart1.set_style(11)     
      
    # ��D2��Ԫ�����ͼ����ƫ�ƣ�  
    worksheet.insert_chart('D2', chart1, {'x_offset': 25, 'y_offset': 10})  
       
    #######################################################################  
    #  
    # ����һ����ͼ������  
    chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})  
       
    # Configure the first series.  
    chart2.add_series({  
        'name':       '=Sheet1!$B$1',  
        'categories': '=Sheet1!$A$2:$A$7',  
        'values':     '=Sheet1!$B$2:$B$7',  
    })  
       
    # Configure second series.  
    chart2.add_series({  
        'name':       '=Sheet1!$C$1',  
        'categories': '=Sheet1!$A$2:$A$7',  
        'values':     '=Sheet1!$C$2:$C$7',  
    })  
       
    # Add a chart title and some axis labels.  
    chart2.set_title ({'name': 'Stacked Chart'})  
    chart2.set_x_axis({'name': 'Test number'})  
    chart2.set_y_axis({'name': 'Sample length (mm)'})  
       
    # Set an Excel chart style.  
    chart2.set_style(12)  
       
    # Insert the chart into the worksheet (with an offset).  
    worksheet.insert_chart('D18', chart2, {'x_offset': 25, 'y_offset': 10})  
       
    #######################################################################  
    #  
    # Create a percentage stacked chart sub-type.  
    #  
    chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})  
       
    # Configure the first series.  
    chart3.add_series({  
        'name':       '=Sheet1!$B$1',  
        'categories': '=Sheet1!$A$2:$A$7',  
        'values':     '=Sheet1!$B$2:$B$7',  
    })  
       
    # Configure second series.  
    chart3.add_series({  
        'name':       '=Sheet1!$C$1',  
        'categories': '=Sheet1!$A$2:$A$7',  
        'values':     '=Sheet1!$C$2:$C$7',  
    })  
       
    # Add a chart title and some axis labels.  
    chart3.set_title ({'name': 'Percent Stacked Chart'})  
    chart3.set_x_axis({'name': 'Test number'})  
    chart3.set_y_axis({'name': 'Sample length (mm)'})  
       
    # Set an Excel chart style.  
    chart3.set_style(13)  
       
    # Insert the chart into the worksheet (with an offset).  
    worksheet.insert_chart('D34', chart3, {'x_offset': 25, 'y_offset': 10})  
    #����Բ��ͼ   
    chart4 = workbook.add_chart({'type':'pie'})  
    #��������  
    data = [  
            ['Pass','Fail','Warn','NT'],  
            [333,11,12,22],  
            ]  
    #д������  
    worksheet.write_row('A51',data[0],bold)  
    worksheet.write_row('A52',data[1])  
       
    chart4.add_series({          
        'name':         '�ӿڲ��Ա���ͼ',  
        'categories': '=Sheet1!$A$51:$D$51',  
        'values':     '=Sheet1!$A$52:$D$52',  
        'points':[  
            {'fill':{'color':'#00CD00'}},  
            {'fill':{'color':'red'}},  
            {'fill':{'color':'yellow'}},  
            {'fill':{'color':'gray'}},  
                  ],  
    })  
    # Add a chart title and some axis labels.  
    chart4.set_title ({'name': '�ӿڲ���ͳ��'})  
    chart4.set_style(3)      
#     chart3.set_y_axis({'name': 'Sample length (mm)'})  
      
    worksheet.insert_chart('E52', chart4, {'x_offset': 25, 'y_offset': 10})  
    workbook.close()  
if __name__ == '__main__':  
#     demo1()  
    charts()  
    print('finished...')  
    pass  