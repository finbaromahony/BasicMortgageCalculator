# BasicMortgageCalculator
Basic Program to display details of a Mortgage in a spreadsheet

##This Project is a work in progress
Initial Working Version is now available

startMonth --sm is not yet implemented

# HOW TO RUN
## Installation
from BasicMortgageCalculator folder

```
pip install setuptools
pip install .
```

## Execution
```
bmc --startMonth <month> --startYear <year> --interest <initial interest rate> \
                --duration <length in years> --amount <value of mortgage>
```

## Result
Result of bmc will be a spreadsheet in the folder in which the command was executed

The usefulness of this is once the spreadsheet is opened you can modify the interest rate at a given month
and year and this value is propagated down the sheet.

Also you can input a value in the Extra Payment cell and this will be removed from the overall total; thus
giving you knowledge beforehand of what your repayment will become if you decide to repay and extra sneaky 10K
onto your variable rate mortgage.
