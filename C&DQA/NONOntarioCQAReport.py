from openpyxl import Workbook
from openpyxl.styles import PatternFill, Side
import gettingFeeCodes
import CQAUtilities

def BCandOtherReport(workbook, CQAREF):
    print 'BCandOtherReport', workbook, CQAREF
    grey_highlight = PatternFill(start_color='E6E6E3', end_color='E6E6E3', fill_type='solid')
    sheet = workbook.get_sheet_by_name("CCME (British Columbia)")
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    array_values = CQAUtilities.OntarioResults(CQAREF)
