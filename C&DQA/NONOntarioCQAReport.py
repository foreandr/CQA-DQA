import os

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Side
import gettingFeeCodes
import CQAUtilities
import Utilities

def BCandOtherReport(workbook, CQAREF):
    print 'BCandOtherReport', workbook, CQAREF
    grey_highlight = PatternFill(start_color='E6E6E3', end_color='E6E6E3', fill_type='solid')
    sheet = workbook.get_sheet_by_name("CCME (British Columbia)")
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    array_values = CQAUtilities.OntarioResults(CQAREF)
    locations_in_excel = Utilities.get_names_and_indexes(sheet)

    # A. ---
    sheet.cell(row=10, column=3).value = array_values[0][2]  # Arsenic
    sheet.cell(row=11, column=3).value = array_values[1][2]  # Cadmium
    sheet.cell(row=12, column=3).value = array_values[2][2]  # chromium  # cobalt
    sheet.cell(row=13, column=3).value = array_values[3][2]
    sheet.cell(row=14, column=3).value = array_values[4][2]
    sheet.cell(row=15, column=3).value = array_values[5][2]
    sheet.cell(row=16, column=3).value = array_values[6][2]
    sheet.cell(row=17, column=3).value = array_values[7][2]
    sheet.cell(row=18, column=3).value = array_values[8][2]
    sheet.cell(row=19, column=3).value = array_values[9][2]
    sheet.cell(row=20, column=3).value = array_values[10][2]

    # B. ---
    sheet.cell(row=25, column=4).value = ""
    sheet.cell(row=26, column=4).value = ""

    # C. ---
    sheet.cell(row=30, column=4).value = ""
    sheet.cell(row=32, column=4).value = ""

    # D. ---
    sheet.cell(row=37, column=4).value = ""
    sheet.cell(row=38, column=4).value = ""

    # E. ---
    sheet.cell(row=43, column=4).value = ""
    sheet.cell(row=44, column=4).value = ""

    # Appendix II -----
    sheet.cell(row=51, column=5).value = ""
    sheet.cell(row=52, column=5).value = ""
    sheet.cell(row=53, column=5).value = ""
    sheet.cell(row=54, column=5).value = ""
    sheet.cell(row=55, column=5).value = ""
    sheet.cell(row=57, column=5).value = ""
    sheet.cell(row=58, column=5).value = ""
    sheet.cell(row=59, column=5).value = ""

    # Appendix III
    sheet.cell(row= 96, column=4).value = ""
    sheet.cell(row= 97, column=4).value = ""
    sheet.cell(row= 98, column=4).value = ""
    sheet.cell(row= 99, column=4).value = ""
    sheet.cell(row=101, column=4).value = ""
    sheet.cell(row=102, column=4).value = ""
    sheet.cell(row=103, column=4).value = ""
    sheet.cell(row=104, column=4).value = ""
    sheet.cell(row=105, column=4).value = ""
    sheet.cell(row=106, column=4).value = ""
    sheet.cell(row=107, column=4).value = ""
    sheet.cell(row=109, column=4).value = "ag index"

    saveLocation = os.path.join(r"C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport", CQAREF)
    filename = saveLocation + "\%sReport.xlsx" % CQAREF
    Workbook.save(workbook, filename)

