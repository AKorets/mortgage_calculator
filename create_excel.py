# -*- coding: utf-8 -*-
"""This module create excel with given parameters to the function create_excel"""
from enum import Enum
from dataclasses import dataclass
import xlsxwriter

class MortgageTypes(Enum):
    """This class enum all supported mortgage types"""
    KALAZ = 'קבועה לא צמודה'
    KVUA = 'קבועה צמודה'
    PRAIM = 'פריים'
    FIVE_YEARS_PERCENT = 'משתנה כל 5 לא צמודה'
    FIVE_YEARS_PERCENT_ZAMUD = 'משתנה כל 5 צמודה'
    GOVERMENT_DIRECT = 'זכאות משרד השיכון'

@dataclass
class Maslul:
    """This is dataclass that contains all information about mortgage parts"""
    mtype:MortgageTypes = MortgageTypes.KVUA
    percentage:float = 4
    month:int = 300
    money:int = 100000

@dataclass
class MainPageInfo:
    """This is dataclass that contains information about excel main page structure"""
    name = 'מחשבון משכנתא'
    main_table_index:int = 4
    maslul_name = 'maslul'
    ribit_mishtana = 0.003

    def get_maslul_name(self, index):
        """get definition of maslul name"""
        return f'{self.maslul_name}{index+1}'

def create_main(main_info, workbook, maslul_arr):
    """create main page in excel"""
    worksheet = workbook.add_worksheet(main_info.name)
    worksheet.right_to_left()

    worksheet.set_column('A:A', 18)
    worksheet.set_column('B:B', 15)

    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    nis_format1 = workbook.add_format({'num_format': '"₪" #,##;[Red]"₪" -#,##'})
    nis_format2 = workbook.add_format({'num_format': '#,##;[Red]-#,##'})

    #header name
    worksheet.write('A1', 'מחשבון משכנתא')

    st = main_info.main_table_index
    #maslul table
    worksheet.write_row('A3', ('מסלול', 'סכום', 'שנים', 'ריבית'))
    for index, maslul in enumerate(maslul_arr):
        worksheet.write_row(f'A{index+st}', (maslul.mtype.value, None, maslul.month/12, None))
        worksheet.write(f'B{index+st}',maslul.money, nis_format1)
        worksheet.write(f'D{index+st}',maslul.percentage/100, percent_fmt)


    worksheet.set_column('F:F', 15)
    worksheet.write('F2','הצג יתרה בשנה:')

    worksheet.write('H2','תשלום חודשי')
    worksheet.write('H3','התחלתי')
    for index, maslul in enumerate(maslul_arr):
        worksheet.write(f'H{index+st}', f'={main_info.get_maslul_name(index)}!D9', nis_format2)

    worksheet.write('I3','ממוצע')
    for index, maslul in enumerate(maslul_arr):
        worksheet.write(f'I{index+st}', f'={main_info.get_maslul_name(index)}!G8', nis_format2)

    worksheet.write('J3','מקסימלי')
    for index, maslul in enumerate(maslul_arr):
        worksheet.write(f'J{index+st}', f'={main_info.get_maslul_name(index)}!G7', nis_format2)

def create_maslul_kalaz_yearly_change(main_info, workbook, worksheet, index, maslul):
    worksheet = create_maslul_kalaz(main_info, workbook, worksheet, index, maslul)

    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    worksheet.write('D3', 'שינוי שנתי')
    worksheet.write('E3', f'{main_info.ribit_mishtana}')

    for i in range(61, 361, 60):
        current_line_xs = 11+1+i
        line_number = 11+i
        worksheet.write(line_number, 3, f'=C3+(E3*((A{current_line_xs}-1)/12))', percent_fmt)

def create_maslul_kalaz(main_info, workbook, worksheet, index, maslul):
    """create page that calculated kalaz"""
    worksheet.right_to_left()
    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    worksheet.set_column('F:F', 18)
    worksheet.set_column('B:B', 13)
    worksheet.set_column('C:C', 13)

    nis_format = workbook.add_format({'num_format': '"₪" #,##0.00;[Red]"₪" -#,##0.00'})
    worksheet.write('A1', maslul.mtype.value)

    worksheet.write('B3', 'ריבית התחלתית')
    worksheet.write('B4', 'סה"כ חודשים')
    worksheet.write('B5', 'סכום')
    worksheet.write('B6', 'מספר תשלומים בשנה')

    InterestYearly = maslul.percentage/100
    worksheet.write('C3', InterestYearly, percent_fmt)

    #workbook.define_name("'%s'!Lsum"%worksheet_name, "='%s'!$H$1"%worksheet_name)

    worksheet.write('C4', maslul.month)

    LSum = maslul.money
    worksheet.write('C5', LSum, nis_format)

    worksheet.write('C6', 12)
    worksheet.write('B9', 'תשלום התחלתי')
    worksheet.write('C9', '=IFERROR(-PMT($C$3/12,$C$4,$C$5,0,0),"")', nis_format)
    worksheet.write('D9', '=ROUNDUP($C$9,2)', nis_format)
    worksheet.write('B10', 'סה"כ')
    worksheet.write('C10', '=SUMIF(B13:B372,">0")', nis_format)

    #header
    worksheet.write_row('A12', ('חודש', 'תשלום חודשי', 'תשלום חודשי מעוגל', 'ריבית',
                                'ריבית', 'ע"ח ריבית', 'ע"ח קרן',  'Extra Payments Here',
                                'יתרת הלוואה', 'תשלומי ריבית מצטברים'))

    worksheet.write('A13', 1)
    worksheet.write('B13', '=$D$9', nis_format)
    worksheet.write('C13', '=$B$13', nis_format)
    worksheet.write('D13', '=C3', percent_fmt)
    worksheet.write('E13', '=D13', percent_fmt)
    worksheet.write('F13', '=ROUNDUP($C$5*$C$3/12,2)', nis_format)
    worksheet.write('G13', '=C13-F13', nis_format)
    worksheet.write('I13', '=$C$5-G13-H13', nis_format)
    worksheet.write('J13', '=F13', nis_format)

    worksheet.write('Q12', 'תשלום חודשי')
    worksheet.write('Q13', '=B13', nis_format)

    for i in range(2,361):
        current_line_xs = 11+1+i
        line_number = 11+i
        if_header = f'IF(I{current_line_xs-1}<=0.01,"",'
        #'=IF(I14<0.02,"",A14+1)'
        worksheet.write(line_number, 0, f'=IF(I{current_line_xs-1}<0.02,"",A{current_line_xs-1}+1)')
        worksheet.write(line_number, 1, f'={if_header}IF(E{current_line_xs}<>E{current_line_xs-1},'+
            f'PMT(E{current_line_xs}/12,$C$4-A{current_line_xs-1},-I{current_line_xs-1}),'+
            f'B{current_line_xs-1}))',
            nis_format)
        worksheet.write(line_number, 2, f'={if_header}IF(B{current_line_xs}>K{current_line_xs-1},'+
                                        f'K{current_line_xs-1},ROUNDUP(B{current_line_xs},2)))',
                                         nis_format)
        worksheet.write(line_number, 3, '', percent_fmt)
        worksheet.write(line_number, 4, f'={if_header}IF(D{current_line_xs}<>"",'+
                                        f'D{current_line_xs},'+
                                        f'E{current_line_xs-1}))', percent_fmt)
        worksheet.write(line_number, 5, f'={if_header}ROUNDUP(I{current_line_xs-1}*'+
                                        f'E{current_line_xs}/12,2))', nis_format)
        worksheet.write(line_number, 6, f'={if_header}C{current_line_xs}-F{current_line_xs})',
                                            nis_format)
        worksheet.write(line_number, 7, 0, nis_format)
        worksheet.write(line_number, 8, f'={if_header}I{current_line_xs-1}-G{current_line_xs}'+
                                        f'-H{current_line_xs})', nis_format)
        worksheet.write(line_number, 9, f'=IF(A{current_line_xs}="","",J{current_line_xs-1}+'+
                                        f'F{current_line_xs})', nis_format)
        worksheet.write(line_number, 10, f'=I{current_line_xs}+(I{current_line_xs}*'+
                                         f'E{current_line_xs+1}/12)', nis_format)
        worksheet.write(line_number, 16, f'=B{current_line_xs}', nis_format)

    worksheet.write('C14', '=IF(I13<=0.01,"",ROUNDUP(B14,2))', nis_format)

    worksheet.write('F7', 'תשלום מקסימלי')
    worksheet.write('F8', 'תשלום ממוצע')
    worksheet.write('F9', 'תשלומי ריבית')

    worksheet.write('G7', '{=MAX(IF(ISERROR(B13:B372),0,B13:B372))}', nis_format)
    worksheet.write('G8', '=IFERROR(AVERAGEIF(B13:B372,">0"),"")', nis_format)
    worksheet.write('G9', '=IF(C9="","",SUMIF(F13:F372,">0"))', nis_format)
    return worksheet

def create_maslul_kvua(main_info, workbook, worksheet, index, maslul):
    """create excel page that calculated kvua"""
    worksheet.right_to_left()
    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    nis_format = workbook.add_format({'num_format': '"₪" #,##0.00;[Red]"₪" -#,##0.00'})

    worksheet.set_column('F:F', 18)
    worksheet.set_column('B:B', 18)

    worksheet.write('A1', maslul.mtype.value)
    worksheet.write('A5', 'סה"כ')
    worksheet.write('B5', '=SUMIF(I13:I372,">0")', nis_format)

    worksheet.write('C1', "ריבית שנתית")
    worksheet.write('C2', "מדד שנתי")
    worksheet.write('C3', "עליית מדד חודשי")

    InterestYearly = maslul.percentage/100
    worksheet.write('D1', InterestYearly, percent_fmt)
    worksheet.write('D2', '0.02')
    worksheet.write('D3', '=RATE(12,0,1,-(1+D2))*100')

    worksheet.write('F1', "סכום הלוואה")
    worksheet.write('F2', "מדד בסיס")
    worksheet.write('F3', "מספר תקופות")
    worksheet.write('F4', "ריבית חודשית")

    LSum = maslul.money
    worksheet.write('H1', LSum, nis_format)
    worksheet.write('H2', 100)

    worksheet.write('H3', maslul.month)
    worksheet.write('H4', '=D1/12', percent_fmt)

    worksheet.write('I1', 'Lsum')
    worksheet.write('I2', 'Base_Index')
    worksheet.write('I4', 'Interest')

    worksheet_name = main_info.get_maslul_name(index)

    workbook.define_name("'%s'!Lsum"%worksheet_name, "='%s'!$H$1"%worksheet_name)
    workbook.define_name("'%s'!Base_Index"%worksheet_name, "='%s'!$H$2"%worksheet_name)
    workbook.define_name("'%s'!Interest"%worksheet_name, "='%s'!$H$4"%worksheet_name)
    #workbook.define_name('Lsum', "=H1")
    #workbook.define_name('Base_Index', '=H2')
    #=Lsum

    worksheet.write('B9', 'תשלום התחלתי')
    #worksheet.write('L1', "התחלתי")
    #worksheet.write('M1', '=IFERROR(IF(I13=0,"",I13),"")')
    worksheet.write('C9', '=IFERROR(IF(I13=0,"",I13),"")', nis_format)
    worksheet.write('D9', '=ROUNDUP($C$9,2)', nis_format)

    worksheet.write('J1', "תשלום ממוצע")
    worksheet.write('J2', "תשלום מקסימלי")
    worksheet.write('J3', 'סה"כ ריבית')
    worksheet.write('J4', 'סה"כ הצמדה')

    #print header
    header_line = 12
    worksheet.write('A%d'%header_line, 'חודש')
    worksheet.write('B%d'%header_line, 'יתרה בתחילת חודש')
    worksheet.write('C%d'%header_line, 'ע"ח קרן')
    worksheet.write('D%d'%header_line, 'ע"ח ריבית')
    worksheet.write('E%d'%header_line, 'סה"כ')
    worksheet.write('F%d'%header_line, 'יתרה בסוף חודש')
    worksheet.write('G%d'%header_line, 'יתרה בסוף חודש כולל מדד')
    worksheet.write('H%d'%header_line, 'מדד בסוף חודש')
    worksheet.write('I%d'%header_line, 'תשלום כולל שינוי מדד')
    worksheet.write('J%d'%header_line, 'ע"ח קרן כולל מדד')
    worksheet.write('K%d'%header_line, 'ע"ח ריבית כולל מדד')

    first_line = 13
    worksheet.write('A%d'%first_line, 1)
    worksheet.write('B%d'%first_line, '=Lsum', nis_format)
    worksheet.write('C%d'%first_line, '=E%d-D%d'%(first_line, first_line), nis_format)
    worksheet.write('D%d'%first_line, '=B%d*Interest'%first_line, nis_format)
    worksheet.write('E%d'%first_line, '=-PMT(Interest,H$3,Lsum)', nis_format)
    worksheet.write('F%d'%first_line, f'=IF(B{first_line}>0.01,B{first_line}-'+
                                       f'C{first_line},0)', nis_format)
    worksheet.write('G%d'%first_line, '=H%d/Base_Index*F%d'%(first_line, first_line), nis_format)

    worksheet.write('H%d'%first_line, '=Base_Index+D$3', nis_format)
    worksheet.write('I%d'%first_line, f'=IF(B{first_line}>0.01,H{first_line}/'+
                                      f'Base_Index*E{first_line},0)', nis_format)
    worksheet.write('J%d'%first_line, f'=IF(B{first_line}>0.01,H{first_line}/'+
                                      f'Base_Index*C{first_line},0)', nis_format)
    worksheet.write('K%d'%first_line, '=H%d/Base_Index*D%d'%(first_line, first_line), nis_format)

    worksheet.write('Q12', 'תשלום חודשי')
    worksheet.write('Q13', '=I13', nis_format)

    for i in range(2,361):
        current_line_xs = 11+1+i
        line_number = 11+i
        worksheet.write(line_number, 0, i)
        worksheet.write(line_number, 1, '=F%d'%(current_line_xs-1), nis_format)
        worksheet.write(line_number, 2, '=E%d-D%d'%(current_line_xs ,current_line_xs), nis_format)
        worksheet.write(line_number, 3, '=B%d*Interest'%current_line_xs, nis_format)
        worksheet.write(line_number, 4, '=-PMT(Interest,H$3,Lsum)', nis_format)
        worksheet.write(line_number, 5, f'=IF(B{current_line_xs}>0.01,B{current_line_xs}-'+
                                        f'C{current_line_xs},0)', nis_format)
        worksheet.write(line_number, 6, f'=H{current_line_xs}/Base_Index*'+
                                        f'F{current_line_xs}', nis_format)
        worksheet.write(line_number, 7, f'=H{current_line_xs-1}+D$3*'+
                                        f'H{current_line_xs-1}/100', nis_format)
        worksheet.write(line_number, 8, f'=IF(B{current_line_xs}>0.01,H{current_line_xs}/'+
                                        f'Base_Index*E{current_line_xs},0)', nis_format)
        worksheet.write(line_number, 9, f'=IF(B{current_line_xs}>0.01,H{current_line_xs}'+
                                        f'/Base_Index*C{current_line_xs},0)', nis_format)
        worksheet.write(line_number, 10, f'=H{current_line_xs}/Base_Index*'+
                                         f'D{current_line_xs}', nis_format)
        worksheet.write(line_number, 16, '=I%d'%(current_line_xs), nis_format)

    #last line
    worksheet.write('A373', 'סה"כ')
    worksheet.write('C373', '=SUM(C13:C372)', nis_format)
    worksheet.write('D373', '=SUM(D13:D372)', nis_format)
    worksheet.write('E373', '=SUM(E13:E372)', nis_format)
    #worksheet.write('I373', '=SUM(I13:I366)', nis_format)

    worksheet.write('F7', 'תשלום מקסימלי')
    worksheet.write('F8', 'תשלום ממוצע')
    worksheet.write('F9', 'תשלומי ריבית')

    worksheet.write('G7', '{=MAX(IF(ISERROR(I13:I372),0,I13:I372))}', nis_format)
    worksheet.write('G8', '=IFERROR(AVERAGEIF(I13:I372,">0"),"")', nis_format)
    worksheet.write('G9', '=IF(C9="","",SUMIF(D13:D372,">0"))', nis_format)

def create_maslul_praim(workbook, worksheet, maslul):
    """create excel page that calculated praim"""
    worksheet.right_to_left()
    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    nis_format = workbook.add_format({'num_format': '"₪" #,##0.00;[Red]"₪" -#,##0.00'})

    worksheet.set_column('F:F', 18)

    worksheet.write('A1', maslul.mtype.value)

    worksheet.write('D3', 'שינוי שנתי')
    worksheet.write('E3', 'טרפז')

    worksheet.write('F1', 'טרפז')
    worksheet.write('F2', 'עלייה 10 שנים עד 5%, קבוע 10 שנים, ירידה 10 שנים לריבית מקורית')
    worksheet.write('F3', '=IF(E3="טרפז",(0.05-C3)/10,"")', percent_fmt)

    worksheet.set_column('B:B', 13)
    worksheet.write('B3', 'ריבית התחלתית')
    worksheet.write('B4', 'סה"כ חודשים')
    worksheet.write('B5', 'סכום')
    worksheet.write('B6', 'מספר תשלומים בשנה')
    worksheet.write('B9', 'תשלום התחלתי')
    worksheet.write('B10','סה"כ')

    InterestYearly = maslul.percentage/100
    worksheet.write('C3', InterestYearly, percent_fmt)

    worksheet.write('C4', maslul.month)

    LSum = maslul.money
    worksheet.write('C5', LSum, nis_format)

    worksheet.write('C6', 12)
    worksheet.write('C9', '=IFERROR(-PMT($C$3/12,$C$4,$C$5,0,0),"")', nis_format)
    worksheet.write('D9', '=ROUNDUP($C$9,2)', nis_format)
    worksheet.write('C10','=SUMIF(B13:B372,">0")', nis_format)

    #header
    worksheet.write_row('A12', ('חודש', 'תשלום חודשי', 'תשלום חודשי מעוגל', 'ריבית',
                                'ריבית', 'ע"ח ריבית', 'ע"ח קרן',  'Extra Payments Here',
                                'יתרת הלוואה', 'תשלומי ריבית מצטברים',' Balance + Interest '))

    worksheet.write('A13', 1)
    worksheet.write('B13', '=$D$9', nis_format)
    worksheet.write('C13', '=$B$13', nis_format)
    worksheet.write('D13', '=C3', percent_fmt)
    worksheet.write('E13', '=D13', percent_fmt)
    worksheet.write('F13', '=ROUNDUP($C$5*$C$3/12,2)', nis_format)
    worksheet.write('G13', '=C13-F13', nis_format)
    worksheet.write('I13', '=$C$5-G13-H13', nis_format)
    worksheet.write('J13', '=F13', nis_format)

    worksheet.write('Q12', 'תשלום חודשי')
    worksheet.write('Q13', '=B13', nis_format)

    for i in range(2,361):
        current_line_xs = 11+1+i
        line_number = 11+i
        if_header = 'IF(I%d<=0.01,"",'%(current_line_xs-1)
        #'=IF(I14<0.02,"",A14+1)'
        worksheet.write(line_number, 0, f'=IF(I{current_line_xs-1}<0.02,"",A{current_line_xs-1}+1)')
        worksheet.write(line_number, 1, f'={if_header}IF(E{current_line_xs}<>E{current_line_xs-1},'+
                                        f'PMT(E{current_line_xs}/12,$C$4-A{current_line_xs-1},-'+
                                        f'I{current_line_xs-1}),B{current_line_xs-1}))', nis_format)
        worksheet.write(line_number, 2, '=%sIF(B%d>K%d,K%d,ROUNDUP(B%d,2)))'%(if_header, current_line_xs, current_line_xs-1, current_line_xs-1, current_line_xs), nis_format)
        worksheet.write(line_number, 3, '', percent_fmt)
        worksheet.write(line_number, 4, '=%sIF(D%d<>"",D%d,E%d))'%(if_header, current_line_xs, current_line_xs, current_line_xs-1), percent_fmt)
        worksheet.write(line_number, 5, '=%sROUNDUP(I%d*E%d/12,2))'%(if_header, current_line_xs-1, current_line_xs), nis_format)
        worksheet.write(line_number, 6, '=%sC%d-F%d)'%(if_header, current_line_xs, current_line_xs), nis_format)
        worksheet.write(line_number, 7, 0, nis_format)
        worksheet.write(line_number, 8, '=%sI%d-G%d-H%d)'%('IF(I%d<=0.01,0,'%(current_line_xs-1), current_line_xs-1, current_line_xs, current_line_xs), nis_format)
        worksheet.write(line_number, 9, '=IF(A%d="","",J%d+F%d)'%(current_line_xs, current_line_xs-1, current_line_xs), nis_format)
        worksheet.write(line_number, 10, '=I%d+(I%d*E%d/12)'%(current_line_xs, current_line_xs, current_line_xs+1), nis_format)
        worksheet.write(line_number, 16, '=B%d'%(current_line_xs), nis_format)

    worksheet.write('C14', '=IF(I13<=0.01,"",ROUNDUP(B14,2))', nis_format)

    for i in range(13,133,12):
        current_line_xs = 11+1+i
        line_number = 11+i

        worksheet.write(line_number, 3, '=IF(E3="טרפז",D%d+F$3,D%d+E$3)'%(current_line_xs-12, current_line_xs-12), percent_fmt)

    for i in range(145,242,12):
        current_line_xs = 11+1+i
        line_number = 11+i

        worksheet.write(line_number, 3, '=IF(E3="טרפז",D133,D%d+E$3)'%(current_line_xs-12), percent_fmt)

    for i in range(241,361,12):
        current_line_xs = 11+1+i
        line_number = 11+i

        worksheet.write(line_number, 3, '=IF(E3="טרפז",D%d-F$3,D%d+E$3)'%(current_line_xs-12, current_line_xs-12), percent_fmt)

    worksheet.write('F7', 'תשלום מקסימלי')
    worksheet.write('F8', 'תשלום ממוצע')
    worksheet.write('F9', 'תשלומי ריבית')

    worksheet.write('G7', '{=MAX(IF(ISERROR(B13:B372),0,B13:B372))}', nis_format)
    worksheet.write('G8', '=IFERROR(AVERAGEIF(B13:B372,">0"),"")', nis_format)
    worksheet.write('G9', '=IF(C9="","",SUMIF(F13:F372,">0"))', nis_format)

def create_maslul_page(main_info, workbook, index, maslul):
    """general definition of each excel pages where calculated maslulim"""
    worksheet = workbook.add_worksheet(main_info.get_maslul_name(index))
    worksheet.right_to_left()
    if maslul.mtype == MortgageTypes.KALAZ:
        create_maslul_kalaz(main_info, workbook, worksheet, index, maslul)
    elif maslul.mtype == MortgageTypes.KVUA:
        create_maslul_kvua(main_info, workbook, worksheet, index, maslul)
    elif maslul.mtype == MortgageTypes.PRAIM:
        create_maslul_praim(workbook, worksheet, maslul)
    elif maslul.mtype == MortgageTypes.FIVE_YEARS_PERCENT:
        create_maslul_kalaz_yearly_change(main_info, workbook, worksheet, index, maslul)
    elif maslul.mtype == MortgageTypes.FIVE_YEARS_PERCENT_ZAMUD:
        create_maslul_kvua(main_info, workbook, worksheet, index, maslul)

def create_excel(excel_path, maslul_arr):
    'main function, that create whole excel bases on given information'
    main_info = MainPageInfo()
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(excel_path)
    create_main(main_info, workbook, maslul_arr)
    for index, maslul in enumerate(maslul_arr):
        create_maslul_page(main_info, workbook, index, maslul)

    workbook.close()
