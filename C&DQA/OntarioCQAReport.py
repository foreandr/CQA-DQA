from openpyxl import Workbook
from openpyxl.styles import PatternFill, Side
import gettingFeeCodes


def OntarioQuebecCQA(workbook, CQAREF):
    print  workbook, CQAREF
    grey_highlight = PatternFill(start_color='E6E6E3', end_color='E6E6E3', fill_type='solid')
    sheet = workbook.get_sheet_by_name("CCME (Ontario)")
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    feecodeList = gettingFeeCodes.gettingfeeCodes('%s' % CQAREF)
    for i in feecodeList:
        print i