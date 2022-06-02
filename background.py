import datetime
import pyautogui
import pywhatkit
import os
from openpyxl import load_workbook


def clear():
    print("\n"*30)


X = datetime.datetime.now()
d = int(X.strftime("%d"))  # date
m = int(X.strftime("%m"))  # month
H = int(X.strftime("%H"))  # hour
M = int(X.strftime("%M"))  # minute

wb = load_workbook("DATABASE FOR BILLING.xlsx")
ws = wb.active
disc = 45

for row in range(1, ws.max_row+1):
    B_day = ws["R"+str(row)].value
    B_mon = ws["S"+str(row)].value
    A_day = ws["T"+str(row)].value
    A_mon = ws["U"+str(row)].value
    if B_day == d and B_mon == m and ws["W"+str(row)].value == "U":
        ph_no = ws["B"+str(row)].value
        name = ws["C"+str(row)].value
        loyalty = ws["V"+str(row)].value
        discount = loyalty + 20
        if discount > 45:
            discount = 45
        message = str("Happy birthday " + str(name) + " May all your wishes come true. Today at this special occasion, We provide you with " + discount + "% OFF on everything SHOP AT ***** ENTERPRISE")
        G = pyautogui.position()
        pywhatkit.sendwhatmsg(ph_no, message, 00, M + 2, 30)
        pyautogui.moveTo(1063, 700)
        pyautogui.click(1063, 700)
        pyautogui.press('enter')
        pyautogui.moveTo(G)
        ws["W"+str(row)].vlaue = "S"
        wb.save("DATABASE FOR BILLING.xlsx")
    if A_day == d and B_mon == m and ws["X"+str(row)].value == "U":
        ph_no = ws["B"+str(row)].value
        name = ws["C"+str(row)].value
        loyalty = ws["V"+str(row)].value
        discount = loyalty + 20
        if discount > 45:
            discount = 45
        message = str("Happy anniversary " + name + " May all your wishes come true. Today at this special occasion, We provide you and your loved one with " + discount + "% OFF on everything SHOP AT ***** ENTERPRISE")
        G = pyautogui.position()
        pywhatkit.sendwhatmsg(ph_no, message, 00, M + 2, 30)
        pyautogui.moveTo(1063, 700)
        pyautogui.click(1063, 700)
        pyautogui.press('enter')
        pyautogui.moveTo(G)
        ws["X"+str(row)].vlaue = "S"
        wb.save("DATABASE FOR BILLING.xlsx")
    if row == ws.max_row:
        break
if d == 1 and H == 0 and M == 1:
    for row in range(4, ws.max_row+1):
        ws["V"+str(row)].value = 0
if H == 0 and M == 0:
    for row in range(1, ws.max_row+1):
        ws["W"+str(row)].value = "U"
        ws["X"+str(row)].value = "U"
        break
print("I AM DONE")
clear()
os.system("main.py")
exit()
