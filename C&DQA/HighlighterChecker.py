import openpyxl
from openpyxl.styles import PatternFill


def try_cast_to_num(value):
    try:
        float(value)
        return value
    except:
        return value


def get_ontario_cqa_constraints_A(sheet):
    highlight = PatternFill(start_color='F3F315', end_color='F3F315', fill_type='solid')
    for i in range(9, 20):
        column4_value = try_cast_to_num(sheet.cell(row=i, column=4).value)
        column6_value = sheet.cell(row=i, column=6).value
        if column4_value > column6_value and column4_value != type(str):
            sheet.cell(row=i, column=4).fill = highlight

    # ---
    column4_value = try_cast_to_num(sheet.cell(row=24, column=4).value)
    column6_value = 2
    if column4_value > column6_value and column4_value != type(str):
        sheet.cell(row=24, column=4).fill = highlight
    # ---
    column4_value = try_cast_to_num(sheet.cell(row=25, column=4).value)
    column6_value = 0.5
    if column4_value > column6_value and column4_value != type(str):
        sheet.cell(row=25, column=4).fill = highlight
    # ---
    column4_value = try_cast_to_num(sheet.cell(row=26, column=4).value)
    column6_value = 0.0
    if column4_value > column6_value and column4_value != type(str):
        sheet.cell(row=26, column=4).fill = highlight
    # ---
    column4_value = try_cast_to_num(sheet.cell(row=28, column=4).value)
    column6_value = 0
    if column4_value > column6_value and column4_value != type(str):
        sheet.cell(row=28, column=4).fill = highlight
