__author__ = 'anatolykonkin'

import xlrd
import email_attach
import os

# calculate currents day of year (name of directory)
import datetime
today = datetime.datetime.now()
day_of_year = (today - datetime.datetime(today.year, 1, 1)).days + 1

dr = str(day_of_year)

# new connection to email
e = email_attach.FetchEmail('breusov@vasko.ru', 'VrZrhvmMw5')
# upload attachments from the last email
e.upload('turulev@taganka.biz', str(day_of_year))

# create new .csv files for 1c
fileList = [os.path.normcase(f) for f in os.listdir(dr)]
print(fileList)
# rebuild all downloaded files
for fn in fileList:
    if 'xls' in fn:
        path = dr + '/' + fn
        print(path)
        workbook = xlrd.open_workbook(path)
        sheet = workbook.sheet_by_index(0)
        df = open(dr + '/' + fn.replace('xlsx', 'csv'), 'w')

        for i in range(13, 1000):
            try:
                data = [sheet.cell_value(i, col) for col in range(sheet.ncols)]
                if data[3] != '':
                    c1 = data[1].encode('utf-8')
                    c2 = data[1].encode('utf-8')
                    c3 = (str(data[3])).encode('utf-8').replace('>', '').rstrip('0').rstrip('.')
                    c4 = ('%.2f' % (data[4],)).rstrip('0').rstrip('.')
                    print(c1)
                    df.write(c1 + ';' + c2 + ';' + c3 + ';' + c4 + '\n')
            except IndexError:
                break
    else: continue

e.send_email('breusov@vasko.ru', './' + str(day_of_year))