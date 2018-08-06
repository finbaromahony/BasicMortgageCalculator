# basic program to show mortgage information
import xlwt

import os
import time
import calendar
import argparse

parser = argparse.ArgumentParser(description='TBD')
parser.add_argument('-sm', '--startMonth', help='start month eg. 04', required=False)
parser.add_argument('-sy', '--startYear', help='start year eg. 2011', required=True)
parser.add_argument('-i', '--interest', help='interest rate', required=True)
parser.add_argument('-d', '--duration', help='duration of mortgage in years', required=True)
parser.add_argument('-a', '--amount', help='initial amount of mortgage', required=True)
args = vars(parser.parse_args())


def create_workbook():
    # Create a new Workbook
    new_book = xlwt.Workbook(encoding="utf-8")
    return new_book


def create_sheet(c_book):
    sheet1 = c_book.add_sheet("Mortgage", cell_overwrite_ok=False)
    return sheet1


def adjust_columns(sheet1):
    for colx in range(3, 17):
        width = 3300
        sheet1.col(colx).width = width


def create_styles():
    font = xlwt.Font()
    font.name = 'Times New Roman'

    currency_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'

    normal = xlwt.XFStyle()
    normal.num_format_str = currency_format

    four = xlwt.Pattern()
    four.pattern = xlwt.Pattern.SOLID_PATTERN
    four.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']
    style_four = xlwt.XFStyle()
    style_four.pattern = four
    style_four.font = font

    five = xlwt.Pattern()
    five.pattern = xlwt.Pattern.SOLID_PATTERN
    five.pattern_fore_colour = xlwt.Style.colour_map['light_green']
    style_five = xlwt.XFStyle()
    style_five.pattern = five
    style_five.font = font
    style_five.num_format_str = currency_format

    six = xlwt.Pattern()
    six.pattern = xlwt.Pattern.SOLID_PATTERN
    six.pattern_fore_colour = xlwt.Style.colour_map['light_yellow']
    style_six = xlwt.XFStyle()
    style_six.pattern = six
    style_six.font = font
    style_six.num_format_str = currency_format

    seven = xlwt.Pattern()
    seven.pattern = xlwt.Pattern.SOLID_PATTERN
    seven.pattern_fore_colour = xlwt.Style.colour_map['rose']
    style_seven = xlwt.XFStyle()
    style_seven.pattern = seven
    style_seven.font = font
    style_seven.num_format_str = currency_format

    eight = xlwt.Pattern()
    eight.pattern = xlwt.Pattern.SOLID_PATTERN
    eight.pattern_fore_colour = 22
    style_eight = xlwt.XFStyle()
    style_eight.pattern = eight
    style_eight.font = font
    style_eight.num_format_str = currency_format

    borders = xlwt.Borders()
    borders.left = xlwt.Borders.DOUBLE
    borders.right = xlwt.Borders.DOUBLE
    borders.top = xlwt.Borders.DOUBLE
    borders.bottom = xlwt.Borders.DOUBLE
    style_eight.borders = borders

    styles = [style_four, style_five, style_six, style_seven, style_eight, normal]
    return styles


def write_skeleton(skel_sheet, s_duration):
    # Write the basic outline to the sheet
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    # ROW1
    for index, col in enumerate(cols):
        if index == 2:
            skel_sheet.write(0, index, "MONTH")
        if 2 < index < 15:
            skel_sheet.write(0, index, calendar.month_abbr[index - 2])
        if index == 15:
            skel_sheet.write(0, index, "total_interest")
        if index == 16:
            skel_sheet.write(0, index, "difference")

    # now for the Present
    skel_sheet.write(3, 0, "Years")
    skel_sheet.write(4, 0, int(s_duration))
    skel_sheet.write(5, 0, "Months")
    skel_sheet.write(6, 0, xlwt.Formula("A5*12"))


def row_two_days(sheet1, row, two_year):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 0 and real_row == 2:
            sheet1.write(row, index, "Starting Amount")
        if index == 2:
            sheet1.write(row, index, "DAYS")
        if 2 < index < 15:
            sheet1.write(row, index, calendar.monthrange(int(two_year), index-2)[1])


def row_three_interest(sheet1, row, three_interest, three_amount):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 0 and real_row == 3:
            sheet1.write(row, index, int(three_amount))
        elif index == 2:
            sheet1.write(row, index, "Interest%")
        elif index == 3:
            if real_row == 3:
                sheet1.write(row, index, float(three_interest))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14])+str(real_row-7)))
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula((cols[index-1]+str(real_row))))


def row_four_months(sheet1, row, style):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row, index, "months", style)
        elif index == 3:
            if real_row == 4:
                sheet1.write(row, index, xlwt.Formula("A" + str(7)), style)
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14]) + str(real_row-7) + str(-1)), style)
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula((cols[index-1] + "" + str(real_row)) + str(-1)), style)


def row_five_payment(sheet1, row, five_year, style, normal):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 1:
            sheet1.write(row, index, five_year)
        elif index == 2:
            sheet1.write(row, index, "Payment", style)
        elif index == 3:
            if real_row == 5:
                sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(real_row-2)+"/100/12,"+cols[index]+"" +
                                                      str(real_row-1)+","+cols[index-3]+""+str(real_row-2)+")"), style)
            else:
                sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(real_row-2)+"/100/12,"+cols[index]+"" +
                                                      str(real_row-1)+","+cols[14]+str(real_row-6)+")"), style)
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(real_row-2)+"/100/12,"+cols[index]+"" +
                                                  str(real_row-1)+","+cols[index-1]+""+str(real_row+1)+")"), style)
        if index == 15:
            sheet1.write(row, index, xlwt.Formula("sum(D" + str(real_row) + ":O" + str(real_row) + ")"), normal)


def row_six_remainder(sheet1, row, style):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row, index, "Remainder", style)
        if index == 3:
            if real_row == 6:
                sheet1.write(row, index, xlwt.Formula("A3+D5+D7-D8"), style)
            else:
#                sheet1.write(row, index, xlwt.Formula((cols[14])+str(real_row-7)), style)
                sheet1.write(row, index, xlwt.Formula((cols[14]) + "" + str(real_row-7) + "+" + col + ""
                             + str(real_row - 1) + "+" + col + "" + str(real_row + 1) + "-" + col +
                             "" + str(real_row + 2)), style)
        if 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula(cols[index - 1] + "" + str(real_row) + "+" + col + ""
                                                  + str(real_row - 1) + "+" + col + "" + str(real_row + 1) + "-" + col +
                                                  "" + str(real_row + 2)), style)


def row_seven_interest(sheet1, row, style, normal):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row, index, "Interest", style)
        if index == 5 or index == 8 or index == 11 or index == 14:
            sheet1.write(row, index, xlwt.Formula("((" + cols[index-1] + "" + str(real_row-1) +
                                                  "*(" + cols[index] + "" + str(real_row-4) +
                                                  "/100))" + "/365)*(" + cols[index] + "" + str(real_row-5) + "+" +
                                                  cols[index-1] + "" + str(real_row-5) + "+" + cols[index-2] + "" +
                                                  str(real_row-5) + ")"), style)
        elif 2 < index < 14:
            sheet1.write(row, index, "", style)
        if index == 15:
            sheet1.write(row, index, xlwt.Formula("sum(D" + str(real_row) + ":O" + str(real_row) + ")"), normal)
        if index == 16:
            sheet1.write(row, index, xlwt.Formula(cols[index-1] + "" + str(real_row) + "+" +
                                                   cols[index-1] + "" + str(real_row-2)), normal)

def row_eight_extra(sheet1, row, style):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    for index, col in enumerate(cols):
        if index == 2:
            sheet.write(row, index, "Extra Pay", style)
        if 2 < index < 15:
            sheet1.write(row, index, "", style)


def main_sheet(sheet1, main_year, main_duration, main_amount, main_style):
    row = 0
    count = int(main_duration) + 1
    while count > 0:
        row += 1
        row_two_days(sheet1, row, main_year)

        row += 1
        row_three_interest(sheet1, row, interest, main_amount)

        row += 1
        row_four_months(sheet1, row, main_style[0])

        row += 1
        row_five_payment(sheet1, row, main_year, main_style[1], main_style[5])

        row += 1
        row_six_remainder(sheet1, row, main_style[2])

        row += 1
        row_seven_interest(sheet1, row, main_style[3], main_style[5])

        row += 1
        row_eight_extra(sheet1, row, main_style[4])

        main_year = int(main_year) + 1
        count = int(count) - 1


def save_and_finish(new_book):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    if not os.path.exists(dir_path + '/spreadsheets'):
        os.mkdir(dir_path + '/spreadsheets')
    new_book.save("./spreadsheets/"+str(int(time.time()))+"_mortgage.xls")
    # -*- coding: utf-8 -*-


year = args['startYear']
interest = args['interest']
duration = args['duration']
amount = args['amount']
startMonth = args['startMonth']
startYear = args['startYear']
book = create_workbook()
sheet = create_sheet(book)
style_list = create_styles()
adjust_columns(sheet)
write_skeleton(sheet, duration)
main_sheet(sheet, year, duration, amount, style_list)
save_and_finish(book)
