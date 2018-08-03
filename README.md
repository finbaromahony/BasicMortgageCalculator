# BasicMortgageCalculator
Basic Program to display details of a Mortgage in a spreadsheet

##This Project is a work in progress
Initial Working Version is now available
Further work required in order to make the 
output in the spreadsheet more readable (colours, fonts etc)

# HOW TO RUN
# dependencies
pip install xlwt
pip install xlrd

# execution
mortgageCalc.py --startMonth <month> --startYear <year> --interest <initial interest rate> \
                --duration <length in years> --amount <value of mortgage>

Result of mortgageCalc will be a spreadsheet in the spreadsheet folder which the script will create

The usefulness of this is once the spreadsheet is opened you can modify the interest rate at a given month
and year and this value is propagated down the sheet.
Also you can input a value in the Extra Payment cell and this will be removed from the overall total thus
giving you knowledge beforehand of what your repayment will become if you decide to repay and extra sneaky 10K
onto your variable rate mortgage.
