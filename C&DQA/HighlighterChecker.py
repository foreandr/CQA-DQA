import openpyxl
from openpyxl.styles import PatternFill


def try_cast_to_num(value):
    if value == '' or len(str(value)) == 0:
        return ''
    try:
        float(value)
        return value
    except:
        return value


def get_ontario_cqa_constraints_A(sheet):
    highlight = PatternFill(start_color='F3F315', end_color='F3F315', fill_type='solid')
    # sheet.cell(row=9, column=4).value = 14 TESTER
    for i in range(9, 20):
        column4_value = try_cast_to_num(sheet.cell(row=i, column=4).value)
        column6_value = sheet.cell(row=i, column=6).value
        if type(column4_value) == unicode:
            continue
        if column4_value > column6_value:
            sheet.cell(row=i, column=4).fill = highlight
            # print('CURRENT TYPE', type(column4_value), column4_value)

    column4_value = try_cast_to_num(sheet.cell(row=24, column=4).value)
    column6_value = 1  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=24, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=25, column=4).value)
    column6_value = 0.5  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=25, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=26, column=4).value)
    column6_value = 0  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=26, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=28, column=4).value)
    column6_value = 0  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=28, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=29, column=4).value)
    column6_value = 0  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=29, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=33, column=4).value)
    column6_value = 4  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=33, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=35, column=4).value)
    column6_value = 400  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value >= column6_value or column4_value == None:
        sheet.cell(row=35, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=40, column=4).value)
    column6_value = 1000  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value >= column6_value:
        sheet.cell(row=40, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=41, column=4).value)
    column6_value = 3  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    try:
        casted_value = float(column4_value)
    except:
        casted_value = column4_value
    if type(casted_value) == float:
        if casted_value >= column6_value:
            sheet.cell(row=41, column=4).fill = highlight
    else:
        print'NOT A NUMBER', (casted_value)

def get_non_ontario_cqa_constraints(sheet):
    highlight = PatternFill(start_color='F3F315', end_color='F3F315', fill_type='solid')
    for i in range(10, 21):
        column4_value = try_cast_to_num(sheet.cell(row=i, column=4).value)
        column6_value = sheet.cell(row=i, column=5).value
        if type(column4_value) == unicode:
            continue
        if column4_value > column6_value:
            sheet.cell(row=i, column=4).fill = highlight
            # print('CURRENT TYPE', type(column4_value), column4_value)

    column4_value = try_cast_to_num(sheet.cell(row=26, column=4).value)
    column6_value = 1  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=26, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=29, column=4).value)
    column6_value = 0  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=29, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=30, column=4).value)
    column6_value = 0  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=30, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=34, column=4).value)
    column6_value = 4  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value:
        sheet.cell(row=34, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=36, column=4).value)
    column6_value = 400  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value or column4_value == None:
        sheet.cell(row=36, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=41, column=4).value)
    column6_value = 400  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    elif column4_value > column6_value or column4_value == None:
        sheet.cell(row=41, column=4).fill = highlight

    column4_value = try_cast_to_num(sheet.cell(row=42, column=4).value)
    column6_value = 3  # COMPARISON VALUE
    if column4_value == 'BDL':
        pass
    try:
        casted_value = float(column4_value)
    except:
        casted_value = column4_value
    if type(casted_value) == float:
        if casted_value >= column6_value:
            sheet.cell(row=42, column=4).fill = highlight
    else:
        print'NOT A NUMBER', (casted_value)