#basic program to show mortgage information
import xlwt
import xlrd

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
#neeed input arguments ssssssssso that wwe caaaaaaaaaan calcuuuulate the number of days
#need to introduce argeparce

def createWorkbook():
    # Create a new Workbook
    book = xlwt.Workbook(encoding="utf-8")
    return book

def createSheet(book):
    sheet1 = book.add_sheet("Mortgage",cell_overwrite_ok=False)
    return sheet1

def writeSkeleton(sheet1, year, interest, duration, amount):
    #Write the basic outline to the sheet
    #get variables from parser
    print(year)
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    #ROW1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(0,index,"MONTH")
        if 2 < index < 15:
            sheet1.write(0,index, calendar.month_abbr[index - 2])
        if index == 15:
            sheet1.write(0,index, "total_interest_paid")

    #now for the Present
    sheet1.write(3,0,"Years")
    sheet1.write(4,0, int(duration))
    sheet1.write(5,0,"Months")
    sheet1.write(6,0, xlwt.Formula("A5*12"))


def rowTwoDays(sheet1, row, year):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    print "rowTwoDays "+str(realrow)
    for index, col in enumerate(cols):
        if index == 0 and realrow == 2:
            sheet1.write(row, index, "Starting Amount")
        if index == 2:
            sheet1.write(row, index, "DAYS")
        if 2 < index < 15:
            sheet1.write(row, index, calendar.monthrange(int(year),index-2)[1])


def rowThreeInterest(sheet1, row, interest):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    for index, col in enumerate(cols):
        if index == 0 and realrow == 3:
            sheet1.write(row, index, int(amount))
        elif index == 2:
            sheet1.write(row, index, "Interest%")
        elif index == 3:
            if realrow == 3:
                sheet1.write(row, index, float(interest))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14])+str(realrow-7)))
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula((cols[index-1]+str(realrow))))


def rowFourMonths(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row,index,"months")
        elif index == 3:
            if realrow == 4:
                sheet1.write(row, index, xlwt.Formula("A"+str(7)))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14])+str(realrow-7)+str(-1)))
        elif 3 < index < 15:
            sheet1.write(row,index, xlwt.Formula((cols[index-1]+""+str(realrow))+str(-1)))


def rowFivePayment(sheet1, row, year):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    for index, col in enumerate(cols):
        if index == 1:
            sheet1.write(row, index, year)
        elif index == 2:
            sheet1.write(row, index, "Payment")
        elif index  == 3:
            if realrow == 5:
                sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[index-3]+""+str(realrow-2)+")"))
            else:
                sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[14]+str(realrow-7)+")"))
        elif 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula("PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[index-1]+""+str(realrow+1)+")"))


def rowSixRemainder(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row,index,"Remainder")
        if index == 3:
            if realrow == 6:
                sheet1.write(row,index,xlwt.Formula("A3+D5+D7-D8"))
            else:
                sheet1.write(row, index, xlwt.Formula((cols[14])+str(realrow-7)))
        if 3 < index < 15:
            sheet1.write(row,index,xlwt.Formula(cols[index-1]+""+str(realrow)+"+"+col+""+str(realrow-1)+"+"+col+""+str(realrow+1)+"-"+col+""+str(realrow+2)))


#TODO: fix
def rowSevenInterest(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row,index,"Interest")
        if index == 5 or index == 8 or index == 11 or index == 14:
            sheet1.write(row,index,xlwt.Formula("(("+
                cols[index-1]+""+str(realrow-1)+
                "*("+cols[index]+""+str(realrow-4)+
                "/100))"+"/365)*("+
                cols[index]+""+str(realrow-5)+"+"+
                cols[index-1]+""+str(realrow-5)+"+"+
                cols[index-2]+""+str(realrow-5)+
                ")"))

#TODO: fix
def rowEightExtra(sheet1, row):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    realrow = row + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet.write(row,index,"Extra Pay")
        if 2< index < 15:
            sheet1.write(row, index, "")

def mainSheet(sheet1, year, duration, amount):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    row = 0
    subtotal = 0
    count = int(duration) + 1
    while (count > 0):
        row += 1
        rowTwoDays(sheet1, row, year)

        row += 1
        rowThreeInterest(sheet1, row, interest)

        row += 1
        rowFourMonths(sheet1, row)

        row += 1
        rowFivePayment(sheet1, row, year)

        row += 1
        rowSixRemainder(sheet1, row)

        row += 1
        rowSevenInterest(sheet1, row)

        row += 1
        rowEightExtra(sheet1, row)

        year = int(year) + 1
        count = int(count) - 1

def saveAndFinish(book):
    book.save("./spreadsheets/"+str(int(time.time()))+"_mortgage.xls")
    # -*- coding: utf-8 -*-

year = args['startYear']
interest = args['interest']
duration = args['duration']
amount = args['amount']
startMonth = args['startMonth']
startYear = args['startYear']
book=createWorkbook()
sheet=createSheet(book)
writeSkeleton(sheet, year, interest, duration, amount)
mainSheet(sheet, year, duration, amount)
saveAndFinish(book)
