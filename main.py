import openpyxl
# import pyautogui
import time
# import pywhatkit
import datetime
import os
from openpyxl import load_workbook, workbook
from openpyxl.utils import get_column_letter

# load workbook
# WB = load_workbook("ZMCH.XLSX") #future update
wb = load_workbook('DATABASE FOR BILLING.xlsx')
# load worksheet
# WS = WB.active #future update
ws = wb.active
srNO = 0
X = 0
K = 0
MRP = 100
while True:
    phno = str(input("input Ph no {format - 91+XXXXXXXXXX}\t"))
    while X == 0:
        for row in range(2, 1048577):
            pch = ws["B"+str(row)].value
            if pch == phno:
                name = ws["C"+str(row)].value
                bday = ws["R"+str(row)].value
                bmon = ws["S"+str(row)].value
                aday = ws["T"+str(row)].value
                amon = ws["U"+str(row)].value
                loyalty = int(ws["V"+str(row)].value) + 0.1
                X = 1
                break
            if pch != phno and row == ws.max_row:
                name = str(input("Enter name \t"))
                bday = int(input("Enter birth date\t"))
                bmon = int(input("Enter birth month\t"))
                aday = int(input("Enter anniversary date\t"))
                amon = int(input("Enter anniversary month\t"))
                loyalty = 0
                X = 1
                break
    for row in range(2, 1048577):
        W = ws["A"+str(row)].value
        if W != row - 1:
            srNO = int(row) - 1
            billno = ws["H" + str(srNO)].value
            break
    X = datetime.datetime.now()
    d = int(X.strftime("%d"))
    m = int(X.strftime("%m"))
    H = int(X.strftime("%H"))
    M = int(X.strftime("%M"))
    MOP = str(input("method of payment\t"))
    noA = int(input("No. of articles\t"))
    Y = 0
    while Y != noA:
        Art_no = int(input("Enter the same article number\t"))
        QNT = int(input("Enter no. of similar article\t"))
        # mrp =
        amount = QNT * MRP
        discount = loyalty
        if d == bday or d == aday:
            if m == bmon or m == amon:
                discount = 20 + discount
        SP = amount - amount * discount
        Y = Y + QNT
        ws.append([srNO, phno, name, d, m, H, M, billno, Art_no, "desc", QNT, "OUM", MRP, amount, discount, SP, MOP, bday, bmon, aday, amon, loyalty, "F"])
    wb.save('DATABASE FOR BILLING.xlsx')
    print("SAVED")
    print("\n\n\n ***THANKS FOR SHOPPING WITH US***\n\n\n")
    os.system('cls')
    time.sleep(3)
    os.system("main.py")
    exit()
    
