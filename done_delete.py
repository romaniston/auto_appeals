import openpyxl

book = openpyxl.open("AppealsCompleted.xlsx")
sheet = book.active
line = sheet[2][0].value = str('')
n = 2

while n != 102:
    sheet[n][0].value = str('')
    n += 1
else:
    book.save('AppealsCompleted.xlsx')