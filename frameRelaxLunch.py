# Form of Schedule
#
import os
import time
import xlsxwriter

# Create
workbook = xlsxwriter.Workbook('./out.xlsx')
worksheet = workbook.add_worksheet('result')
worksheet_raw_copy = workbook.add_worksheet('raw_copy')
workbook.sheetnames['result']== worksheet
workbook.sheetnames['raw_copy'] == worksheet_raw_copy
# Row and column default
worksheet.set_column('A:A', 5)
worksheet.set_column('B:C', 15)
worksheet.set_column('D:J', 30)
worksheet.set_default_row(45)
worksheet.set_zoom(70)
#Wirte form
#title
format_title = workbook.add_format({
    'bold' :1, 
    'border' : 1,
    'font': 'Noto Sans Arabic',
    'size' : 30,
    'align' : 'left',
    'valign' : 'vcenter',
    'border_color': '#F2B3E8',
    'bg_color':'#78A66A',
    'font_color': '#C3F3B4' })
worksheet.merge_range('B1:J2', 'CLASS-S-CHEDULE', format_title)
#row main
format_cell = workbook.add_format({
    'bold' :0, 
    'border' : 1,
    'font': 'SF Pro Display',
    'size' : 15,
    'align' : 'center',
    'valign' : 'vcenter',
    'border_color': '#F2B3E8',
    'bg_color':'#78A66A',
    'font_color': '#C3F3B4' })
datarow= ('LESSON', 'TIME', 'MONDAY','TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY','SUNDAY')
worksheet.write_row('B3', datarow , format_cell)

#format all cell core
format_cell_core = workbook.add_format({
    'bold' :0, 
    'border' : 1,
    'font': 'SF Pro Display',
    'size' : 15,
    'align' : 'center',
    'valign' : 'vcenter',
    'text_wrap': True,
    'border_color': '#F2B3E8',
    'right_color' :  '#FFFFFF',
    'left_color': '#FFFFFF',
    'font_color': '#000000' })
for row_num in range (3,19):
    for col_num in range (3,10):
        worksheet.write_row(row_num, col_num,'=CHAR(10)',format_cell_core)

#Timeline on day //have relax time on
datalesson=('1','2','3','4','5',"RELAX",'6','7','8','9','10', '11','12','13')
worksheet.write_column('B4', datalesson , format_cell)
for row_num in range(0, 6):
    formula = '=CONCATENATE((B'+str(row_num+4)+'+6),":00")'
    worksheet.write_formula(row_num+3, 2, formula,format_cell, 0)
for row_num in range(6, 15):
    formula = '=CONCATENATE((B'+str(row_num+4)+'+7),":00")'
    worksheet.write_formula(row_num+3, 2, formula,format_cell, 0)
worksheet.write('C9','FREE',format_cell)
#Remove border
format_border = workbook.add_format({ 'border': 0})
worksheet.merge_range('A1:A17', '', format_border)
worksheet.merge_range('A18:Z100', '', format_border)
worksheet.merge_range('K1:Z17', '', format_border)
worksheet.write('D4:K18', '', format_border)

worksheet.write(3,3,'=CHAR(10)',format_cell_core)

workbook.close()
#Check end
#print ('===Successfull frame.py===')
