# basic program to show mortgage information
import xlwt

import os
import time
import calendar
import argparse

parser = argparse.ArgumentParser(description='TBD')
parser.add_argument('-sm','--startMonth', help='start month eg. 04', required=True)
parser.add_argument('-sy','--startYear', help='start year eg. 2011', required=True)
parser.add_argument('-i','--interest', help='interest rate', required=True)
parser.add_argument('-d','--duration', help='duration of mortgage in years', required=True)
parser.add_argument('-a','--amount', help='initial amount of mortgage', required=True)
args = vars(parser.parse_args())


def create_workbook():
    # Create a new Workbook
    new_book = xlwt.Workbook(encoding="utf-8")
    return new_book


def create_sheet(book):
    sheet1 = book.add_sheet("Mortgage",cell_overwrite_ok=False)
    return sheet1


def write_skeleton(sheet1, year, duration):
    # Write the basic outline to the sheet
    print(year)
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    # ROW1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(0, index, "MONTH")
        if 2 < index < 15:
            sheet1.write(0, index, calendar.month_abbr[index - 2])
        if index == 15:
            sheet1.write(0, index, "total_interest_paid")

    # now for the Present
    sheet1.write(3, 0, "Years")
    sheet1.write(4, 0, int(duration))
    sheet1.write(5, 0,"Months")
    sheet1.write(6, 0, xlwt.Formula("A5*12"))


def row_two_days(sheet1, row, year):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 0 and real_row == 2:
            sheet1.write(row, index, "Starting Amount")
        if index == 2:
            sheet1.write(row, index, "DAYS")
        if 2 < index < 15:
            sheet1.write(row, index, calendar.monthrange(int(year), index-2)[1])


def row_three_interest(sheet1, row, interest, amount):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 0 and real_row == 3:
            sheet1.write(row, index, int(amount))
        elif index == 2:
            sheet1.write(row, index, "Interest%")
        elif index == 3:
            if real_row == 3:
                sheet1.write(row, index, float(interest))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14])+str(real_row-7)))
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula((cols[index-1]+str(real_row))))


def row_four_months(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row, index, "months")
        elif index == 3:
            if real_row == 4:
                sheet1.write(row, index, xlwt.Formula("A" + str(7)))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14]) + str(real_row-7) + str(-1)))
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula((cols[index-1] + "" + str(real_row)) + str(-1)))


def row_five_payment(sheet1, row, year):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 1:
            sheet1.write(row, index, year)
        elif index == 2:
            sheet1.write(row, index, "Payment")
        elif index == 3:
            if real_row == 5:
                sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(real_row-2)+"/100/12,"+cols[index]+""+
                                                      str(real_row-1)+","+cols[index-3]+""+str(real_row-2)+")"))
            else:
                sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(real_row-2)+"/100/12,"+cols[index]+""+
                                                      str(real_row-1)+","+cols[14]+str(real_row-6)+")"))
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(real_row-2)+"/100/12,"+cols[index]+""+
                                                  str(real_row-1)+","+cols[index-1]+""+str(real_row+1)+")"))


def row_six_remainder(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row, index, "Remainder")
        if index == 3:
            if real_row == 6:
                sheet1.write(row, index, xlwt.Formula("A3+D5+D7-D8"))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14])+str(real_row-7)))
        if 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula(cols[index - 1] + "" + str(real_row) + "+" + col + ""
                                                  + str(real_row - 1) + "+" + col + "" + str(real_row + 1) + "-" + col +
                                                  "" + str(real_row + 2)))


def row_seven_interest(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    real_row = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row, index, "Interest")
        if index == 5 or index == 8 or index == 11 or index == 14:
            sheet1.write(row, index, xlwt.Formula("((" + cols[index-1] + "" + str(real_row-1) +
                                                  "*(" + cols[index] + "" + str(real_row-4) +
                                                  "/100))" + "/365)*(" + cols[index] + "" + str(real_row-5) + "+" +
                                                  cols[index-1] + "" + str(real_row-5) + "+" + cols[index-2] + "" +
                                                  str(real_row-5) + ")"))


def row_eight_extra(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    for index, col in enumerate(cols):
        if index == 2:
            sheet.write(row, index, "Extra Pay")
        if 2 < index < 15:
            sheet1.write(row, index, "")


def main_sheet(sheet1, year, duration, amount):
    row = 0
    count = int(duration) + 1
    while count > 0:
        row += 1
        row_two_days(sheet1, row, year)

        row += 1
        row_three_interest(sheet1, row, interest, amount)

        row += 1
        row_four_months(sheet1, row)

        row += 1
        row_five_payment(sheet1, row, year)

        row += 1
        row_six_remainder(sheet1, row)

        row += 1
        row_seven_interest(sheet1, row)

        row += 1
        row_eight_extra(sheet1, row)

        year = int(year) + 1
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
write_skeleton(sheet, year, duration)
main_sheet(sheet, year, duration, amount)
save_and_finish(book)
