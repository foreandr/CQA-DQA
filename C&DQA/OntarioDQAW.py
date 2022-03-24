import os

from openpyxl.styles import PatternFill, Border, Side
from openpyxl import Workbook
import Colors
import Utilities


def OntarioPrintDQA(Workbook, CQARef):
    # Sets the color of the highlight/fill to highlight the failed values
    highlight = PatternFill(start_color='F3F315', end_color='F3F315', fill_type='solid')
    grey_highlight = PatternFill(start_color='E6E6E3', end_color='E6E6E3', fill_type='solid')
    sheet = Workbook.get_sheet_by_name("Ontario CFIA", )
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    import gettingFeeCodes

    feecodeList = gettingFeeCodes.gettingfeeCodes('%s' % CQARef)
    print Colors.bcolors.OKGREEN, '\nFEECODELIST', Colors.bcolors.ENDC
    for i in feecodeList:
        print i
    # Some missing, but actually turned out pretty good
    locations_in_excel = Utilities.grab_excel_locations()
    for row in locations_in_excel:
        # print row
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
            # print(value, 'not a float, entering into dict as string')
            typcasted_dict[key] = value
    # A
    sheet.cell(row=17, column=4).value = typcasted_dict['Molybdenum (Mb)']

    # B.
    sheet.cell(row=29, column=4).value = feecodeList[35][3]

    # C
    sheet.cell(row=27, column=4).value = feecodeList[36][3]
    sheet.cell(row=28, column=4).value = feecodeList[37][3]

    # C Pathogens
    sheet.cell(row=36, column=4).value = feecodeList[1][3]
    sheet.cell(row=35, column=4).value = feecodeList[3][3]

    # E Physical Parameter
    # sheet.cell(row=53, column=4).value = feecodeList[7][3]
    # sheet.cell(row=54, column=4).value = feecodeList[45][3]
    # sheet.cell(row=55, column=4).value = feecodeList[6][3]

    # Minimum Agricultural Values
    sheet.cell(row=53, column=4).value = feecodeList[11][3]
    sheet.cell(row=54, column=4).value = float(feecodeList[38][3]) * typcasted_dict[
        'Dry Matter'] / 100  # this is total phophate not available
    sheet.cell(row=55, column=4).value = float(feecodeList[39][3]) * typcasted_dict['Dry Matter'] / 100

    # Agricultural End-Use
    sheet.cell(row=63, column=4).value = feecodeList[9][3]
    sheet.cell(row=64, column=4).value = feecodeList[44][3]  # can't seem to find
    sheet.cell(row=65, column=4).value = Utilities.getTotalOrganicMatter(CQARef)
    sheet.cell(row=66, column=4).value = feecodeList[45][3]
    sheet.cell(row=67, column=4).value = feecodeList[6][3]

    # Fertilizer Equivalent Materials
    sheet.cell(row=76, column=4).value = feecodeList[28][3]  # SODIUM
    # sheet.cell(row=76, column=4).value = feecodeList[40][3]
    # sheet.cell(row=78, column=4).value = feecodeList[41][3]
    # sheet.cell(row=80, column=4).value = feecodeList[43][3]
    # sheet.cell(row=81, column=4).value = feecodeList[33][3]
    # sheet.cell(row=82, column=4).value = feecodeList[16][3]
    # sheet.cell(row=83, column=4).value = feecodeList[12][3]
    # sheet.cell(row=84, column=4).value = feecodeList[21][3]
    # sheet.cell(row=85, column=4).value = feecodeList[22][3]
    # sheet.cell(row=86, column=4).value = feecodeList[26][3]
    # sheet.cell(row=87, column=4).value = feecodeList[27][3]
    # sheet.cell(row=88, column=4).value = feecodeList[34][3]
    # ----------------------------------------------------------------    print '\n'
    # for key, value in typcasted_dict.iteritems():
    #   print key, value

    '''FOR MULTIPLICATION PURPOSES'''
    # print('\nDRY MATTER : ')
    # print(typcasted_dict['Dry Matter'])
    calc_value = typcasted_dict['Nitrogen Total (N)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=71, column=4).value = round(calc_value, 1)

    calc_value = typcasted_dict['Nitrate Nitrogen NO3-N'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=73, column=4).value = calc_value

    calc_value = typcasted_dict['Total Phosphate (P as P2O5)'] * typcasted_dict['Dry Matter'] / 100
    num_x = calc_value / 10000 * 2.29  # LOOKS SAME BUT IS ACTUALLY DIFFERENT
    sheet.cell(row=74, column=4).value = num_x

    calc_value = typcasted_dict['Total Potash (K as K2O)'] * typcasted_dict['Dry Matter'] / 100
    num_x = calc_value / 10000 * 1.21  # LOOKS SAME BUT IS ACTUALLY DIFFERENT
    sheet.cell(row=75, column=4).value = num_x

    calc_value = typcasted_dict['Available Sodium (Na)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    # print('sodium calc value', calc_value)
    sheet.cell(row=76, column=4).value = calc_value

    calc_value = typcasted_dict['Sodium'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=77, column=4).value = calc_value

    calc_value = typcasted_dict['Total Available (Mg)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=78, column=4).value = calc_value

    calc_value = typcasted_dict['Total Magnesium (Mg)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=79, column=4).value = calc_value

    calc_value = typcasted_dict['Total available (Ca)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=80, column=4).value = calc_value

    calc_value = typcasted_dict['Total Calcium (Ca)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=81, column=4).value = calc_value

    calc_value = typcasted_dict['Available (S}'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=82, column=4).value = round(calc_value, 1)

    calc_value = typcasted_dict['Total Sulfur (S)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=83, column=4).value = calc_value

    calc_value = typcasted_dict['Boron (B)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=84, column=4).value = calc_value

    calc_value = typcasted_dict['Chloride'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=85, column=4).value = calc_value

    calc_value = typcasted_dict['Copper (Cu)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=86, column=4).value = calc_value

    calc_value = typcasted_dict['Iron (Fe)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=87, column=4).value = calc_value

    calc_value = typcasted_dict['Manganese (Mn)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=88, column=4).value = calc_value

    calc_value = typcasted_dict['Molybdenum (Mb)'] * typcasted_dict['Dry Matter'] / 100 / 10000  # small but seems right
    sheet.cell(row=89, column=4).value = calc_value

    calc_value = typcasted_dict['Zinc (Zn)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=90, column=4).value = calc_value

    # --------------------------------------------
    from CQAUtilities import DQA_ONT_FORMATTING
    DQA_ONT_FORMATTING(sheet)
    # -----------------------------------------------------------------
    # BORDER ALIGNMENT
    """
    from Utilities import FixFormatting
    #A intro table MADNESS FORMATTING
    for i in range(7, 8):
        border = Border(left=thick, top=thick, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)
        border = Border(left=thick, right=thick)
        FixFormatting(sheet, 'B8:I8', border)
        border = Border(left=thick, right=thick)
        FixFormatting(sheet, 'B9:I9', border)
        border = Border(left=thick, right=thick)
        FixFormatting(sheet, 'B10:I10', border)
        border = Border(top=thin)
        FixFormatting(sheet, 'E8:H8', border)
        border = Border(bottom=thin)
        FixFormatting(sheet, 'D9:H9', border)
        border = Border(left=thin)
        FixFormatting(sheet, 'D8:D8', border)
        border = Border(left=thin, right=thin)
        FixFormatting(sheet, 'D9:D9', border)
        border = Border(left=thin, right=thin)
        FixFormatting(sheet, 'F9:F9', border)
        border = Border(left=thin, right=thin)
        FixFormatting(sheet, 'H9:H9', border)

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
        border = Border(left=thin)
        FixFormatting(sheet, 'D35:D35', border)
        border = Border(left=thin)
        FixFormatting(sheet, 'I35:I35', border)

        border = Border(top=thin)
        FixFormatting(sheet, 'E35:H35', border)

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

        border = Border(left=thin)
        FixFormatting(sheet, 'A55:A55', border)

    for i in range(56, 60):
        border = Border(right=thin, left=thin)
        FixFormatting(sheet, 'D62:D62', border)

        border = Border(top=thin)
        FixFormatting(sheet, 'F62:I62', border)
        border = Border(left=thick)
        FixFormatting(sheet, 'A62:A62', border)

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

        sheet['F62'].fill = grey_highlight

        border = Border(right=thick, top=thick, bottom=thin)
        FixFormatting(sheet, 'B64:I64', border)

        border = Border(right=thick, bottom=thick)
        FixFormatting(sheet, 'B68:I68', border)

        border = Border(right=thick)
        FixFormatting(sheet, 'B63:I63', border)

        border = Border(right=thick)
        FixFormatting(sheet, 'B69:I69', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    for i in range(71, 90):
        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B70:I70', border)

        border = Border(top=thin, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'B89:I89', border)
    """
    # ROUNDING THINGS
    """
    print('\nPRINTING Excel values')
    for i in range(70, 90):
        print(i, sheet.cell(row=i, column=6).value, sheet.cell(row=i, column=7).value, sheet.cell(row=i, column=8).value, sheet.cell(row=i, column=9).value)
        column6_formula = Utilities.add_round_to_excel_formula(sheet.cell(row=i, column=6).value)
        column7_formula = Utilities.add_round_to_excel_formula(sheet.cell(row=i, column=7).value)
        column8_formula = Utilities.add_round_to_excel_formula(sheet.cell(row=i, column=8).value)
        column9_formula = Utilities.add_round_to_excel_formula(sheet.cell(row=i, column=9).value)

        sheet.cell(row=i, column=6).value = column6_formula
        sheet.cell(row=i, column=7).value = column7_formula
        sheet.cell(row=i, column=8).value = column8_formula
        sheet.cell(row=i, column=9).value = column9_formula

        # sheet.write_formula('F70', new_formula)
    """
    # CENTERING THINGS

    from openpyxl.styles import Alignment
    from openpyxl.styles import Font
    for i in range(1, 100):
        # print(sheet.cell(row=i, column=4).value)
        # sheet.cell(row=i, column=4).value.alignment = Alignment(horizontal='center')
        current_cell = sheet['D%d' % i]
        current_cell.alignment = Alignment(horizontal='center')
        current_cell.font = Font(bold=True, name='Franklin Gothic Book')

        if i in range(70, 90):
            current_cell = sheet['F%d' % i]
            current_cell.alignment = Alignment(horizontal='center')
            current_cell = sheet['G%d' % i]
            current_cell.alignment = Alignment(horizontal='center')
            current_cell = sheet['H%d' % i]
            current_cell.alignment = Alignment(horizontal='center')
            current_cell = sheet['I%d' % i]
            current_cell.alignment = Alignment(horizontal='center')

    if float(sheet['D88'].value) < 1.0:  # should proably turn this into a function, way easier and more maliable
        sheet['D88'] = 'BDL'
        sheet['F88'] = 'N/A'
        sheet['G88'] = 'N/A'
        sheet['H88'] = 'N/A'
        sheet['I88'] = 'N/A'

    # putting in the images------------------------------------
    from openpyxl.drawing.image import Image
    os.chdir(r'C:\CQA\FULL CQA - DQA\C&DQA\Photos')

    img = Image('al.jpg')
    sheet.add_image(img, 'B1')
    img = Image('Digestate-logo.png')
    sheet.add_image(img, 'I1')

    img = Image('al.jpg')
    sheet.add_image(img, 'A48')
    img = Image('Digestate-logo.png')
    sheet.add_image(img, 'I48')

    # ----------------------------------------------------------------
    saveLocation = os.path.join(r"C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport",
                                CQARef)
    Workbook.save(saveLocation + "\%sReport.xlsx" % (CQARef))
