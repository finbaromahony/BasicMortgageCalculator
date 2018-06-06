#!/usr/bin/python
#calculate the interest

#accepts interest rate and other variables to determine the interest
import numpy


def pmt(interest, periods, value):
    #(r(PV))/(1-(1+r)^-n))=P
    #P=payment
    #r=Interest rate per period
    #n=number of periods
    #PV=Present Value
    interest=((interest/100)/12)
    return numpy.pmt(interest,periods,value)
