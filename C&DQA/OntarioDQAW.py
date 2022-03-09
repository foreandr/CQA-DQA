import os

from openpyxl.styles import PatternFill, Border, Side

import Utilities

ENVDict = {
    "1": "Arsenic",
    "2": "Cadmium",
    "3": 'Chromium',
    "4": 'Cobalt',
    "5": 'Copper',
    '6': 'Lead',
    "7": 'Mercury',
    '8': 'Molybdenum',
    '9': 'Nickel',
    '10': 'Selenium',
    '11': 'Zinc',
    '12': 'Total FM > 25 mm',
    '13': 'Total sharps > 2.8 mm*',
    '14': 'Total sharps > 12.5 mm',
    '15': 'Respiration-mgCO2-C/g OM/day',
    '16': 'E. coli',
    '17': 'Salmonella spp.',
    '18': 'Total Organic Matter',
    '19': 'Moisture',
    '20': 'Total Organic Matter @ 550 deg C',
    '22': 'C:N Ratio',
    '28': 'Total Solids (as received)',
    '30': 'Bulk Density (As Recieved)',
    '33': 'Ammonia (NH3/NH4-N)',
    '34': 'Total Phosphorus (As P205)',
    '35': 'Total Potassium (as K20)',
    '36': 'Calcium',
    '37': 'Magnesium',
    '38': 'Sulphur',
    '39': 'Total FM > 2.8 mm*',
    '40': 'Total plastics > 2.8 mm*'
}


def OntarioPrintDQA(Workbook, CQARef):
    # Sets the color of the highlight/fill to highlight the failed values
    highlight = PatternFill(start_color='F3F315', end_color='F3F315', fill_type='solid')
    sheet = Workbook.get_sheet_by_name("Ontario CFIA", )
    # sheet = Workbook.get_sheet_by_name("K:/2022Ontario/Student2022/ANDRE-CQA 2 report system/Templates/Ontario DQA -W")
    # newDict = Utilities.associate_nums_with_values(finalResult, ENVDict)

    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    # print 'new dict demo -------\n'
    # print ENVDict
    # print finalResult
    # print newDict
    # print 'new dict demo -------\n'

    import gettingFeeCodes

    feecodeList = gettingFeeCodes.gettingfeeCodes('%s' % CQARef)
    print '\n'
    for i in feecodeList:
        print i
    # Some missing, but actually turned out pretty good
    locations_in_excel = Utilities.open_report_csv_INDEXLIST()
    print '\n'
    for row in locations_in_excel:
        print row

        current_index = row[0]
        current_name = row[1]
        for i in feecodeList:
            if i[1] == current_name:
                # print current_index, current_name, i[1]
                # print 'yes'
                sheet.cell(row=current_index, column=4).value = i[3]

    dict_names_values = {}
    for i in feecodeList:
        dict_names_values[i[1]] = (i[3])
    typcasted_dict = {}
    print '\n'
    for key, value in dict_names_values.iteritems():
        try:
            typcasted_dict[key] = float(value)
        except:
            print(value, 'not a float, entering into dict as string')
            typcasted_dict[key] = value
    # A
    sheet.cell(row=18, column=4).value = typcasted_dict['Molybdenum (Mb)']

    # B. Foreign Matter [21, 24]
    sheet.cell(row=28, column=4).value = feecodeList[36][3]
    sheet.cell(row=29, column=4).value = feecodeList[37][3]
    sheet.cell(row=30, column=4).value = feecodeList[35][3]

    # D Pathogens
    sheet.cell(row=37, column=4).value = feecodeList[1][3]
    sheet.cell(row=38, column=4).value = feecodeList[3][3]

    # E Physical Parameter
    sheet.cell(row=52, column=4).value = feecodeList[7][3]
    sheet.cell(row=53, column=4).value = feecodeList[45][3]
    sheet.cell(row=54, column=4).value = feecodeList[6][3]

    # Minimum Agricultural Values
    sheet.cell(row=57, column=4).value = feecodeList[11][3]
    sheet.cell(row=58, column=4).value = float(feecodeList[38][3]) * typcasted_dict['Dry Matter'] / 100  # this is total phophate not available
    sheet.cell(row=59, column=4).value = float(feecodeList[39][3]) * typcasted_dict['Dry Matter'] / 100

    # Agricultural End-Use
    sheet.cell(row=67, column=4).value = feecodeList[9][3]
    sheet.cell(row=68, column=4).value = feecodeList[44][3]  # can't seem to find

    # Fertilizer Equivalent Materials
    sheet.cell(row=76, column=4).value = feecodeList[28][3]
    sheet.cell(row=77, column=4).value = feecodeList[40][3]
    sheet.cell(row=79, column=4).value = feecodeList[41][3]

    sheet.cell(row=81, column=4).value = feecodeList[43][3]
    sheet.cell(row=82, column=4).value = feecodeList[33][3]

    sheet.cell(row=83, column=4).value = feecodeList[16][3]
    sheet.cell(row=84, column=4).value = feecodeList[12][3]
    sheet.cell(row=85, column=4).value = feecodeList[21][3]
    sheet.cell(row=86, column=4).value = feecodeList[22][3]
    sheet.cell(row=87, column=4).value = feecodeList[26][3]
    sheet.cell(row=88, column=4).value = feecodeList[27][3]
    sheet.cell(row=89, column=4).value = feecodeList[34][3]

    # ----------------------------------------------------------------
    print '\n'
    for key, value in typcasted_dict.iteritems():
        print key, value

    '''FOR MULTIPLICATION PURPOSES'''
    print('\nDRY MATTER : ')
    print(typcasted_dict['Dry Matter'])
    calc_value = typcasted_dict['Nitrogen Total (N)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=70, column=4).value = round(calc_value, 1)

    calc_value = typcasted_dict['Nitrate Nitrogen NO3-N'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=72, column=4).value = calc_value

    calc_value = typcasted_dict['Total Phosphate (P as P2O5)'] * typcasted_dict['Dry Matter'] / 100
    num_x = calc_value / 10000 * 2.29  # LOOKS SAME BUT IS ACTUALLY DIFFERENT
    sheet.cell(row=73, column=4).value = num_x

    calc_value = typcasted_dict['Total Potash (K as K2O)'] * typcasted_dict['Dry Matter'] / 100
    num_x = calc_value / 10000 * 1.21  # LOOKS SAME BUT IS ACTUALLY DIFFERENT
    sheet.cell(row=74, column=4).value = num_x

    calc_value = typcasted_dict['Available Sodium (Na)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    print('sodium calc value', calc_value)
    sheet.cell(row=75, column=4).value = calc_value

    calc_value = typcasted_dict['Sodium'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=76, column=4).value = calc_value

    calc_value = typcasted_dict['Total Available (Mg)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=77, column=4).value = calc_value

    calc_value = typcasted_dict['Total Magnesium (Mg)'] * typcasted_dict[
        'Dry Matter'] / 100  # oNLY WORKS IF I DONT DO THE SECOND DIVISION?

    sheet.cell(row=78, column=4).value = calc_value

    calc_value = typcasted_dict['Total available (Ca)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    print 'avail calcium value', calc_value
    sheet.cell(row=79, column=4).value = calc_value

    calc_value = typcasted_dict['Total Calcium (Ca)'] * typcasted_dict['Dry Matter'] / 100  # ONLY WOKRS IF I DONT DO SECOND DIVISON
    print 'total calc', calc_value
    sheet.cell(row=80, column=4).value = calc_value

    calc_value = typcasted_dict['Available (S}'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=81, column=4).value = round(calc_value, 1)

    calc_value = typcasted_dict['Total Sulfur (S)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=82, column=4).value = calc_value

    calc_value = typcasted_dict['Boron (B)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=83, column=4).value = calc_value

    calc_value = typcasted_dict['Chloride'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=84, column=4).value = calc_value

    calc_value = typcasted_dict['Copper (Cu)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=85, column=4).value = calc_value

    calc_value = typcasted_dict['Iron (Fe)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=86, column=4).value = calc_value

    calc_value = typcasted_dict['Manganese (Mn)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=87, column=4).value = calc_value

    calc_value = typcasted_dict['Molybdenum (Mb)'] * typcasted_dict['Dry Matter'] / 100 / 10000  # small but seems right
    print('molynum', calc_value)
    sheet.cell(row=88, column=4).value = calc_value

    calc_value = typcasted_dict['Zinc (Zn)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=89, column=4).value = calc_value

    # -----------------------------------------------------------------
    # BORDER ALIGNMENT
    from Utilities import FixFormatting

    for i in range(7, 11):
        border = Border(left=thick, bottom=thick, top=thick, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    for i in range(11, 21):
        border = Border(bottom=thin, top=thin)
        FixFormatting(sheet, 'B%d:H%d' % (i, i), border)

        border = Border(bottom=thin, top=thick)
        FixFormatting(sheet, 'B11:H11', border)

        border = Border(bottom=thick, top=thin)
        FixFormatting(sheet, 'B21:H21', border)

    for i in range(28, 31):
        border = Border(bottom=thick, top=thick, right=thick)
        FixFormatting(sheet, 'B27:I27', border)

        border = Border(bottom=thin, top=thick, right=thick)
        FixFormatting(sheet, 'B28:I28', border)

        border = Border(bottom=thin, top=thin, right=thick)
        FixFormatting(sheet, 'B29:I29', border)

        border = Border(bottom=thick, top=thin, right=thick)
        FixFormatting(sheet, 'B30:I30', border)

    for i in range(34, 39):
        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B34:I34', border)

        border = Border(left=thick, right=thick)
        FixFormatting(sheet, 'B35:I35', border)

        border = Border(left=thick, bottom=thin, top=thick, right=thick)
        FixFormatting(sheet, 'B36:I36', border)

        border = Border(bottom=thin, top=thin, right=thick)
        FixFormatting(sheet, 'B37:I37', border)

        border = Border(bottom=thick, top=thin, right=thick)
        FixFormatting(sheet, 'B38:I38', border)

    for i in range(51, 55):
        border = Border(top=thick, right=thick, bottom=thick)
        FixFormatting(sheet, 'B51:I51', border)

        border = Border(right=thick, bottom=thin)
        FixFormatting(sheet, 'B52:I52', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'B53:I53', border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'B54:I54', border)

    for i in range(56, 60):
        border = Border(top=thick, right=thick, bottom=thick)
        FixFormatting(sheet, 'B56:I56', border)

        border = Border(right=thick, bottom=thin)
        FixFormatting(sheet, 'B57:I57', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'B58:I58', border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'B59:I59', border)

    for i in range(65, 68):
        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B61:I61', border)

        border = Border(right=thick, bottom=thick)
        FixFormatting(sheet, 'B62:I62', border)

        border = Border(right=thick, top=thick, bottom=thin)
        FixFormatting(sheet, 'B64:I64', border)

        border = Border(right=thick, bottom=thick)
        FixFormatting(sheet, 'B68:I68', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    for i in range(71, 88):
        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B70:I70', border)

        border = Border(top=thin, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'B89:I89', border)

    # ----------------------------------------------------------------
    saveLocation = os.path.join(r"C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport",
                                CQARef)
    Workbook.save(saveLocation + "\%sReport.xlsx" % (CQARef))
