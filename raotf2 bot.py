import praw
import openpyxl as xl
import time

index = 0
delay = 3
filename = input('filename (*.xslx): ')
sheetname = input('Sheet Name: ')


wb = xl.load_workbook(filename)
sheet = wb[sheetname]


names = []
for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1):
    for cell in row:
        if cell.value is not None:
            names.append(cell.value)

subjects = []
for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=2, max_col=2):
    for cell in row:
        if cell.value is not None:
            subjects.append(cell.value)

messages = []
for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=3, max_col=3):
    for cell in row:
        if cell.value is not None:
            messages.append(cell.value)

# print(names)
# print(subjects)
# print(messages)

reddit = praw.Reddit(client_id='-_jp3B32nE2jxw',
                     client_secret='qfo3JrQGI0vcm7s4W2u9qrsQDoI',
                     username='BotTesting11',
                     password='BotPassword:123',
                     user_agent='user agent'
                     )
while True:
    index += 1;
    time.sleep(delay)
    reddit.redditor(names[index]).message(subjects[index], messages[index])
    if index >= len(names) - 1:
        break