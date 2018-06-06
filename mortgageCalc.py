#basic program to show mortgage information
import xlwt
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
    sheet1 = book.add_sheet("Mortgage")
    return sheet1

def writeSkeleton(sheet1, year, interest, duration, amount):
    #Write the basic outline to the sheet
    #get variables from parser
    print(year)
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    #ROW0
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(0,index,"MONTH")
        if 2 < index < 15:
            sheet1.write(0,index, calendar.month_abbr[index - 2])
        if index == 15:
            sheet1.write(0,index, "total_interest_paid")

    #ROW1
    for index, col in enumerate(cols):
        if index == 0:
            sheet1.write(1,index, "Starting Amount")
        if index == 2:
            sheet1.write(1,index, "DAYS")
        if 2 < index < 15:
            sheet1.write(1,index, calendar.monthrange(int(year),index-2)[1])
            print(calendar.monthrange(2011,index-2)[1])
            #sheet1.write(1,index, calendar.monthrange(2011,index-2)[1])

    #ROW3
    for index, col in enumerate(cols):
        if index == 0:
            sheet1.write(2,index, int(amount))
        if index == 2:
            sheet1.write(2,index, "Interest%")
        if 2 < index < 15:
            sheet1.write(2, index, float(interest))

    #now for the Present
    sheet1.write(3,0,"Years")
    sheet1.write(4,0, int(duration))
    sheet1.write(5,0,"Months")
    sheet1.write(6,0, xlwt.Formula("A5*12"))

def testSheet():
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]

    row = 3
    realrow = 4
    row = row + 1
    realrow = realrow + 1
    subtotal = 0
    for index, col in enumerate(cols):
        if index == 2:
            print ("current index is " + str(index))
            print ("current col is " + col)
            print ("index less 1 is " + str(index-1))
            #print ("col less 1 is " + col-1)
            print ("other way for col is " + cols[index])
            print ("other way for col less 1 is " + cols[index-1])
            print ("current row is "+str(row)+" written as "+str(row+1)+ "  in real sheet")
            print ("real row is "+str(realrow))

    for index, col in enumerate(cols):
        if index == 3:
            print("=PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[index-3]+""+str(realrow-2)+")")
        if 3 < index < 15:
            #sheet1.write(row, index, xlwt.Formula("=PMT("+cols[index]+""+row-2+"/100/12,"+cols[index]+""+row-1+","+cols[index-3]+""+row-2+")"))
            print("=PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[index-1]+""+str(realrow+1)+")")
    print("=SUM(D"+str(realrow)+":O"+str(realrow))

    row = row + 1
    realrow = realrow + 1

    for index, col in enumerate(cols):
        if index == 3:
            print("=A3+D5+D7-D8")
        if 3 < index < 15:
            print("="+cols[index-1]+""+str(realrow)+"+"+col+""+str(realrow-1)+"+"+col+""+str(realrow+1)+"-"+col+""+str(realrow+2))

def mainSheet(sheet1, year, duration, amount):
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    #row0=monthnames
    #row1=days in the month
    #row2=interest rate
    #row3=Months
    #row4=payments (and current year in row B)
    #row5=remainder
    #row6=Interest
    #row7=extrapayments
    #row8=DAYS
    #row9=interest rate
    #repeat
    row = 3
    realrow =  4
    subtotal = 0

    # the first set needs to be independent as it sets the initial values
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row,index,"months")
        if index == 3:
            sheet1.write(row, index, xlwt.Formula("=A"+str(7)))
        if 3 < index < 15:
            sheet1.write(row,index, xlwt.Forula("="+col+""+row-1))
    row = row + 1
    realrow = realrow + 1
    for index, col in enumerate(cols):
        if index == 1:
            sheet1.write(row, index, year)
        if index == 2:
            sheet1.write(row, index, "Payment")
        if index  == 3:
            sheet1.write(row, index, xlwt.Formula("=PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[index-3]+""+str(realrow-2)+")"))
        if 3 < index < 15:
            sheet1.write(row, index, xlwt.Formula("=PMT("+cols[index]+""+str(realrow-2)+"/100/12,"+cols[index]+""+str(realrow-1)+","+cols[index-1]+""+str(realrow+1)+")"))
    sheet1.write(row,15, xlwt.Formula("=SUM(D"+str(realrow)+":O"+str(realrow)))
    row = row + 1
    realrow = realrow + 1
    for index, col in enumerate(cols):
        if index == 2:
            sheet1.write(row,index,xlwt.Formula("Remainder"))
        if index == 3:
            sheet1.write(row,index,xlwt.Forumla("=A3+D5+D7-D8"))
        if 3 < index < 15:
            print("="+cols[index-1]+""+str(realrow)+"+"+col+""+str(realrow-1)+"+"+col+""+str(realrow+1)+"-"+col+""+str(realrow+2))

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
testSheet()
mainSheet(sheet, year, duration, amount)
saveAndFinish(book)
