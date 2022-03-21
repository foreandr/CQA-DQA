import os

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Side
import gettingFeeCodes
import CQAUtilities
import Utilities


def OntarioQuebecCQA(workbook, CQAREF):
    print  'ONTARIO /QEUEBEC REPORT', workbook, CQAREF
    grey_highlight = PatternFill(start_color='E6E6E3', end_color='E6E6E3', fill_type='solid')
    sheet = workbook.get_sheet_by_name("CCME (Ontario)")
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")
    # sprint('CULMN-NUM, NAME, VALUE, ROW-INDX ')
    array_values = CQAUtilities.OntarioResults(CQAREF)
    item_dict = Utilities.getValuesForAGIndex(CQAREF)
    # I THINK PUTTING THEM IN MANUALLY IS JUST EASIER / SIMPLER TO READ, SORRY FUTURE CODERS

    # A.
    sheet.cell(row=9, column=4).value = array_values[0][2]  # Arsenic
    sheet.cell(row=10, column=4).value = array_values[1][2]  # Cadmium
    sheet.cell(row=11, column=4).value = array_values[3][2]  # chromium
    sheet.cell(row=12, column=4).value = array_values[4][2]  # cobalt
    sheet.cell(row=13, column=4).value = array_values[5][2]
    sheet.cell(row=14, column=4).value = array_values[6][2]
    sheet.cell(row=15, column=4).value = array_values[7][2]
    sheet.cell(row=16, column=4).value = array_values[8][2]
    sheet.cell(row=17, column=4).value = array_values[9][2]
    sheet.cell(row=18, column=4).value = array_values[10][2]
    sheet.cell(row=19, column=4).value = array_values[11][2]

    # B.
    sheet.cell(row=24, column=4).value = array_values[29][2]
    sheet.cell(row=25, column=4).value = array_values[30][2]
    sheet.cell(row=26, column=4).value = array_values[11][2]
    sheet.cell(row=28, column=4).value = array_values[12][2]
    sheet.cell(row=29, column=4).value = array_values[13][2]

    # C.
    sheet.cell(row=33, column=4).value = array_values[14][2]
    #sheet.cell(row=34, column=4).value = array_values[][]

    # D.
    sheet.cell(row=40, column=4).value = array_values[15][2]
    sheet.cell(row=41, column=4).value = array_values[16][2]

    # E.
    sheet.cell(row=46, column=6).value = array_values[17][2]
    sheet.cell(row=47, column=6).value = array_values[18][2]

    pe_m3_dict = CQAUtilities.getOtherResults(CQAREF)


    sheet.cell(row=53, column=6).value = array_values[19][2]
    sheet.cell(row=54, column=6).value = array_values[20][2]
    sheet.cell(row=55, column=6).value = 'N/A'
    sheet.cell(row=56, column=6).value = pe_m3_dict['salt'] # salt
    sheet.cell(row=57, column=6).value = pe_m3_dict['perna_m3'] #perna
    # Major nutrients
    sheet.cell(row=59, column=6).value = pe_m3_dict['perk_m3'] #perk
    sheet.cell(row=60, column=6).value = pe_m3_dict['permg_m3'] #perma
    sheet.cell(row=61, column=6).value = pe_m3_dict['perca_m3'] #perca

    # APENDIX 3
    sheet.cell(row=100, column=4).value = CQAUtilities.get_dry_matter(CQAREF)
    sheet.cell(row=101, column=4).value = item_dict['PH']
    sheet.cell(row=102, column=4).value = array_values[22][2]
    sheet.cell(row=103, column=4).value = array_values[20][2]

    # FERTILIZER
    Nitrogen = Utilities.getNitrogen(CQAREF)
    sheet.cell(row=105, column=4).value = Nitrogen # Nitrogen
    sheet.cell(row=106, column=4).value = array_values[23][2]
    sheet.cell(row=107, column=4).value = array_values[24][2]
    sheet.cell(row=108, column=4).value = array_values[25][2]
    sheet.cell(row=109, column=4).value = array_values[26][2]
    sheet.cell(row=110, column=4).value = array_values[27][2]
    sheet.cell(row=111, column=4).value = array_values[28][2]


    item_dict = Utilities.getValuesForAGIndex(CQAREF)

    Nitrogen = float(Utilities.removePercentSign(Utilities.getNitrogen(CQAREF)))  # stand in for real value
    Phosphorus = float(Utilities.removePercentSign(array_values[24][2]))
    Potassium = float(Utilities.removePercentSign(array_values[25][2]))
    Sodium = float(item_dict['NA'])
    DryMatter = 10
    Chloride = float(item_dict['CL'])
    # print('dude wtf', Chloride, type(Chloride), type(float(Chloride)))

    print('PHOSPHORUS:', Phosphorus)
    print('Potassium:', Potassium)
    print('Nitrogen:', Nitrogen)
    print('sodium:', Sodium)
    print('DryMatter:', DryMatter)
    print('Chloride:', Chloride)
    value1 = (Nitrogen + Phosphorus + Potassium)
    value2 = (Sodium * (DryMatter / 100)) + (Chloride / 10000)
    # print 'LEFTSIDE:', value1
    # print 'RIGHTSIDE:', value2
    ag_index = value1 / value2
    sheet.cell(row=113, column=4).value = ag_index



    # AGINDEX----------------------------------------------------------

    # ----------------------------------------------------------------
    saveLocation = os.path.join(r"C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport",
                                CQAREF)

    filename = saveLocation + "\%sReport.xlsx" % CQAREF
    Workbook.save(workbook, filename)

    # template path | C:\CQA\FULL CQA - DQA\C&DQA\Templates\TEMPLATE ON WRITTEN
