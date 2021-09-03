import re
import sys, os
import urllib3
import requests
import json
import xlrd
import xlwt

recordTypeCode = []
transactionCode = []
bankRoutingNumber = []
accountNumber = []

ACH_Test_Data_File = xlrd.open_workbook(r"C:\Users\jhuin\Documents\Python Scripts\ACH_Test_Data.xlsx")

entryDetail_Read = ACH_Test_Data_File.sheet_by_name('Entry Detail Record')

entryDetail_Headers = entryDetail_Read.row(0)


for b in range(len(entryDetail_Headers)):
    if entryDetail_Headers[b].value == "Record Type Code":
        recordTypeCode = entryDetail_Read.col_values(b)[1:]
    if entryDetail_Headers[b].value == "Transaction Code":
        transactionCode = entryDetail_Read.col_values(b)[1:]
    if entryDetail_Headers[b].value == "Bank Routing Number":
        bankRoutingNumber = entryDetail_Read.col_values(b)[1:]
    if entryDetail_Headers[b].value == "Account Number":
        accountNumber = entryDetail_Read.col_values(b)[1:]
    if entryDetail_Headers[b].value == "Amount":
        amount = entryDetail_Read.col_values(b)[1:]

transitRoutingNumber = 0
accountNumber_RDFI = ""
individualIdNumber = ""

taxId = ["999000001","999000002","999000003","999000004","999000005","999000006","999000007","999000008"]
name = ["JACOB HUINKER","JACOB HUINKER","JACOB HUINKER","JACOB HUINKER","JACOB HUINKER","JACOB HUINKER","JACOB HUINKER","JACOB HUINKER"]
blankSpace = "  "
addendaRecordIndicator = "0"
eastWestBank = "32207038"
traceNumber = "0000001"

individualIdNumber = taxId
while len(individualIdNumber) < 15:
    individualIdNumber += " "

while len(traceNumber) < 7:
    traceNumber = "0" + traceNumber

ach_File = open(r'Sample_ACH_File.txt','w')

for a in range(len(recordTypeCode)):
    #Checks whether routing number is 8 or 9 and calculates check digit accordingly
    if len(bankRoutingNumber[a]) == 9:
        transitRoutingNumber = bankRoutingNumber[a][0:8]
        checkDigit = bankRoutingNumber[a][8]
    else:
        transitRoutingNumber = bankRoutingNumber[a]
        routingSum = (int(transitRoutingNumber[0]) * 3) + (int(transitRoutingNumber[1]) * 7) + (int(transitRoutingNumber[2] * 1) + (int(transitRoutingNumber[3] * 3))) + (int(transitRoutingNumber[4]) * 7) + (int(transitRoutingNumber[5] * 1) + (int(transitRoutingNumber[6] * 3))) + (int(transitRoutingNumber[7]) * 7)
        roundingMultiple = routingSum
        while (roundingMultiple % 10 != 0):
            roundingMultiple += 1
        checkDigit = roundingMultiple - routingSum

    accountNumber_RDFI = accountNumber[a]
    while len(accountNumber_RDFI) < 17:
        accountNumber_RDFI += " "

    if amount[a][len(amount[a])-3] == ".":
        transDollarAmount = "".join(digit for digit in amount[a] if digit.isalnum())
        while len(transDollarAmount) < 10:
            transDollarAmount = "0" + transDollarAmount
    else:
        transDollarAmount = amount[a]
        while len(transDollarAmount) < 10:
            transDollarAmount = "0" + transDollarAmount

    individualIdNumber = taxId[a]
    while len(individualIdNumber) < 15:
        individualIdNumber += " "

    nameField = name[a]
    while (len(nameField)) < 22:
        nameField += " "

    

    ach_File.write("%s%s%s%s%s%s%s%s%s%s%s%s\n" % (recordTypeCode[a],transactionCode[a],transitRoutingNumber,checkDigit,accountNumber_RDFI,transDollarAmount,individualIdNumber,nameField,blankSpace,addendaRecordIndicator,eastWestBank,traceNumber))

ach_File.close()