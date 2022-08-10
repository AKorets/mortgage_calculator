# -*- coding: utf-8 -*-
from enum import Enum
from dataclasses import dataclass
import xlsxwriter

class MortgageTypes(Enum):
    Kalaz = 'קבועה לא צמודה'
    Kvua = 'קבועה צמודה'
    Praim = 'פריים'
    FiveYearsPercent = 'משתנה כל 5 לא צמודה'
    FiveYearsPercentZamud = 'משתנה כל 5 צמודה'
    GovermentDirect = 'זכאות משרד השיכון'
    
@dataclass
class Maslul:
    mtype:MortgageTypes = MortgageTypes.Kvua
    percentage:float = 4
    month:int = 300
    money:int = 100000
    
def create_main(workbook, maslulArr):
    worksheet = workbook.add_worksheet('Main')
    worksheet.right_to_left()
    
    worksheet.set_column('A:A', 18)
    worksheet.set_column('B:B', 15)
    
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    nis_format1 = workbook.add_format({'num_format': '"₪" #,##;[Red]"₪" -#,##'})
        
    #header name
    worksheet.write('A1', 'מחשבון משכנתא')
    
    st = 4 # start table index
    #maslul table
    worksheet.write_row('A3', ('מסלול', 'סכום', 'שנים', 'ריבית'))
    for index, maslul in enumerate(maslulArr):
        worksheet.write_row(f'A{index+st}', (maslul.mtype.value, None, maslul.month/12, None))
        worksheet.write(f'B{index+st}',maslul.money, nis_format1)
        worksheet.write(f'D{index+st}',maslul.percentage/100, percent_fmt)
        

    pass
    
def create_excel(excelPath, maslulArr):
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(excelPath)
    create_main(workbook, maslulArr)
    
    
    workbook.close()
    pass