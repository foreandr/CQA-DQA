import os

from openpyxl.styles import Side, Border

import Colors
import Utilities


def CFIAPrintDQA(Workbook, CQARef):
    print(Workbook, CQARef)
    sheet = Workbook.get_sheet_by_name("CCME CFIA", )
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    import gettingFeeCodes

    feecodeList = gettingFeeCodes.gettingfeeCodes('%s' % CQARef)
    print '\n'

    #for i in feecodeList:
    #    print i

    locations_in_excel = Utilities.get_names_and_indexes(sheet)
    list_for_entering = []
    for i in locations_in_excel:
        print Colors.bcolors.HEADER + str(i) + Colors.bcolors.ENDC
        for j in feecodeList:
            if i[1] == j[1]:
                # print i[0], i[1], j[1], j[3]
                temp_list = [i[0], j[1], j[3]]
                list_for_entering.append(temp_list)
            #else:
            #    print(i[1])
    #print('printing list for entering')
    #for i in list_for_entering:
    #    print i

    for i in feecodeList:
        print Colors.bcolors.OKGREEN + str(i) + Colors.bcolors.ENDC

    for i in range(1, 100):
        for j in list_for_entering:
            if i == j[0]:
                #print i, j
                sheet.cell(row=i, column=4).value = j[2]

    #---------- CREATING CALC DICT
    dict_names_values = {}
    for i in feecodeList:
        dict_names_values[i[1]] = (i[3])
    typcasted_dict = {}
    for key, value in dict_names_values.iteritems():
        try:
            typcasted_dict[key] = float(value)
        except:
            #print(value, 'not a float, entering into dict as string')
            typcasted_dict[key] = value

    for key, value in typcasted_dict.items():
        print key, value

    dry_matter = typcasted_dict['Dry Matter']
    # print dry_matter
    #---------
    # A
    sheet.cell(row=18, column=4).value = feecodeList[27][3]

    # B. Foreign Matter
    sheet.cell(row=26, column=4).value = feecodeList[36][3]
    sheet.cell(row=27, column=4).value = feecodeList[37][3]
    sheet.cell(row=28, column=4).value = feecodeList[35][3]

    # D Pathogens
    sheet.cell(row=33, column=4).value = feecodeList[3][3]

    # E Physical Parameter DOESNT EXIST LOL
    #sheet.cell(row=38, column=4).value = feecodeList[7][3]
    #sheet.cell(row=39, column=4).value = feecodeList[45][3]
    #sheet.cell(row=40, column=4).value = feecodeList[6][3]

    # Minimum Agricultural Values
    sheet.cell(row=51, column=4).value = feecodeList[11][3]
    sheet.cell(row=52, column=4).value = float(feecodeList[38][3]) * typcasted_dict['Dry Matter'] / 100# dry matter division and multiplication
    sheet.cell(row=53, column=4).value = float(feecodeList[39][3]) * typcasted_dict['Dry Matter'] / 100# dry matter division

    # Agricultural End-Use 1
    sheet.cell(row=61, column=4).value = feecodeList[9][3]
    sheet.cell(row=63, column=4).value = feecodeList[7][3]
    sheet.cell(row=64, column=4).value = feecodeList[45][3]
    sheet.cell(row=65, column=4).value = feecodeList[6][3]




    # Agricultural End-Use 2
    #print 'typcasted_dict'
    #print typcasted_dict
    # ---- calculations
    calc_value = typcasted_dict['Nitrogen Total (N)'] * dry_matter / 100
    sheet.cell(row=69, column=4).value = round(calc_value, 1)

    calc_value = typcasted_dict['Nitrate Nitrogen NO3-N'] * dry_matter / 100
    sheet.cell(row=71, column=4).value = calc_value

    calc_value = typcasted_dict['Total Phosphate (P as P2O5)'] * dry_matter / 100
    num_x = calc_value / 10000 * 2.29  # LOOKS SAME BUT IS ACTUALLY DIFFERENT
    sheet.cell(row=72, column=4).value = num_x

    calc_value = typcasted_dict['Total Potash (K as K2O)'] * dry_matter / 100
    num_x = calc_value / 10000 * 1.21  # LOOKS SAME BUT IS ACTUALLY DIFFERENT
    sheet.cell(row=73, column=4).value = num_x

    calc_value = typcasted_dict['Available Sodium (Na)'] * dry_matter / 100 / 10000
    sheet.cell(row=74, column=4).value = calc_value

    calc_value = typcasted_dict['Sodium'] * dry_matter / 100 / 10000
    sheet.cell(row=75, column=4).value = calc_value

    calc_value = typcasted_dict['Total Available (Mg)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=76, column=4).value = calc_value

    calc_value = typcasted_dict['Total Magnesium (Mg)'] * typcasted_dict['Dry Matter'] / 100  # oNLY WORKS IF I DONT DO THE SECOND DIVISION?
    sheet.cell(row=77, column=4).value = calc_value

    calc_value = typcasted_dict['Total available (Ca)'] * typcasted_dict['Dry Matter'] / 100 / 10000
    sheet.cell(row=78, column=4).value = calc_value

    calc_value = typcasted_dict['Total Calcium (Ca)'] * typcasted_dict['Dry Matter'] / 100  # ONLY WOKRS IF I DONT DO SECOND DIVISON
    sheet.cell(row=79, column=4).value = calc_value

    calc_value = typcasted_dict['Available (S}'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=80, column=4).value = round(calc_value, 1)

    calc_value = typcasted_dict['Total Sulfur (S)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=81, column=4).value = calc_value

    calc_value = typcasted_dict['Boron (B)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=82, column=4).value = calc_value

    calc_value = typcasted_dict['Chloride'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=83, column=4).value = calc_value

    calc_value = typcasted_dict['Copper (Cu)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=84, column=4).value = calc_value

    calc_value = typcasted_dict['Iron (Fe)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=85, column=4).value = calc_value

    calc_value = typcasted_dict['Manganese (Mn)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=86, column=4).value = calc_value

    calc_value = typcasted_dict['Molybdenum (Mb)'] * typcasted_dict['Dry Matter'] / 100 / 10000  # small but seems right
    sheet.cell(row=87, column=4).value = calc_value

    calc_value = typcasted_dict['Zinc (Zn)'] * typcasted_dict['Dry Matter'] / 100
    sheet.cell(row=88, column=4).value = calc_value


    # -----------------------------------------------------------------
    from Utilities import FixFormatting
    # BORDER ALIGNMENT
    '''
    
    for i in range(7, 11):
        border = Border(left=thick, bottom=thick, top=thick, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

        border = Border(left=thin, right=thin)
        FixFormatting(sheet, 'D8:D8', border)
        border = Border(left=thin, right=thin)
        FixFormatting(sheet, 'D9:D9', border)
        border = Border(right=thin)
        FixFormatting(sheet, 'H9:H9', border)
        border = Border(bottom=thick)
        FixFormatting(sheet, 'B10:H10', border)

    for i in range(11, 21):
        border = Border(bottom=thin, top=thin)
        FixFormatting(sheet, 'B%d:H%d' % (i, i), border)

        border = Border(bottom=thick, top=thin)
        FixFormatting(sheet, 'B21:H21', border)

    for i in range(25, 28):
        border = Border(bottom=thick, top=thick, right=thick)
        FixFormatting(sheet, 'B25:I25', border)

        border = Border(bottom=thin, top=thick, right=thick)
        FixFormatting(sheet, 'B26:I26', border)

        border = Border(bottom=thin, top=thin, right=thick)
        FixFormatting(sheet, 'B27:I27', border)

        border = Border(bottom=thick, top=thin, right=thick)
        FixFormatting(sheet, 'B28:I28', border)

    for i in range(34, 39):

        border = Border(top=thick, right=thick, bottom=thick)
        FixFormatting(sheet, 'B31:I31', border)

        border = Border(left=thick, right=thick)
        FixFormatting(sheet, 'B32:I32', border)

        border = Border(bottom=thick, top=thin, right=thick)
        FixFormatting(sheet, 'B33:I33', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'B38:I38', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'B39:I39', border)

        border = Border(bottom=thick, top=thin, right=thick)
        FixFormatting(sheet, 'B40:I40', border)

    for i in range(51, 55):
        border = Border(top=thick, right=thick, bottom=thick)
        FixFormatting(sheet, 'A50:I50', border)

        border = Border(left=thin, right=thin)
        FixFormatting(sheet, 'D56:D56', border)
        border = Border(left=thick)
        FixFormatting(sheet, 'A56:A56', border)

        border = Border(top=thin)
        FixFormatting(sheet, 'F56:I56', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'A51:I51', border)

        border = Border(bottom=thin, right=thick)
        FixFormatting(sheet, 'A52:I52', border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'A53:I53', border)

    for i in range(59, 62):
        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B55:I55', border)

        border = Border(right=thick, bottom=thick)
        FixFormatting(sheet, 'B56:I56', border)

        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B58:I58', border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'B62:I62', border)

        border = Border(right=thick)
        FixFormatting(sheet, 'I63:I63', border)

        border = Border(right=thick)
        FixFormatting(sheet, 'I57:I57', border)

        border = Border(bottom=thin, top=thin, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    for i in range(65, 83):
        border = Border(bottom=thin, top=thin, right=thick)
        FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

        border = Border(top=thick, right=thick)
        FixFormatting(sheet, 'B64:I64', border)

        border = Border(bottom=thick, right=thick)
        FixFormatting(sheet, 'B83:I83', border)
    #--- changging bel locations to NA
    if float(sheet['D82'].value) < 1.0: # should proably turn this into a function, way easier and more maliable
        sheet['D82'] = 'BDL'
        sheet['F82'] = 'N/A'
        sheet['G82'] = 'N/A'
        sheet['H82'] = 'N/A'
        sheet['I82'] = 'N/A'
    '''
    #CENTERING THINGS
    from openpyxl.styles import Alignment
    from openpyxl.styles import Font
    for i in range(1, 100):
        # print(sheet.cell(row=i, column=4).value)
        # sheet.cell(row=i, column=4).value.alignment = Alignment(horizontal='center')
        current_cell = sheet['D%d'%i]
        current_cell.alignment = Alignment(horizontal='center')
        current_cell.font = Font(bold=True, name='Franklin Gothic Book')


    # putting in the images------------------------------------
    from openpyxl.drawing.image import Image
    os.chdir(r'C:\CQA\FULL CQA - DQA\C&DQA\Photos')
    img = Image('al.jpg')
    sheet.add_image(img, 'B2')
    img = Image('Digestate-logo.png')
    sheet.add_image(img, 'H2')

    img = Image('al.jpg')
    sheet.add_image(img, 'A47')
    img = Image('Digestate-logo.png')
    sheet.add_image(img, 'H47')
    # ----------------------------------------------------------------
    saveLocation = os.path.join(r"C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport",
                                CQARef)
    Workbook.save(saveLocation + "\%sReport.xlsx" % (CQARef))
    pass
