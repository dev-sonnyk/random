import xlrd
import csv as c

workbook = xlrd.open_workbook('PMTask.xlsx')

sheet1 = workbook.sheet_by_index(0)
sheet2 = workbook.sheet_by_index(1)

tickets = []
'''
for i in range(1, sheet1.nrows) :
    if sheet1.row(i)[0].value not in tickets : tickets.append(sheet1.row(i)[0].value)

for i in range(1, sheet2.nrows) :
    print(sheet2.row(i)[0]) if sheet2.row(i)[0].value not in tickets else tickets.remove(sheet2.row(i)[0].value)

print(len(tickets))

if len(tickets) != 0 :
    for t in tickets : print(t)
'''

for i in range(1, sheet1.nrows) :
    p = sheet1.row(i)[1].value
    pt = sheet1.row(i)[5].value
    tickets.append(p) if pt == '' else tickets.append(pt)

with open('difference.csv', 'w', newline='') as csvfile :
    #print('\nNot Found in SM9 from Tableau\n')
    csvfile.truncate()
    writer = c.writer(csvfile)

    for i in range(1, sheet2.nrows) :
        if not sheet2.row(i)[1].value[:6] == "2017/12" :
            p = sheet2.row(i)[0].value
            pt = sheet2.row(i)[2].value
            if pt == '' :
                writer.writerow([str(p)]) if p not in tickets else tickets.remove(p)
            else :
                writer.writerow([str(pt)]) if pt not in tickets else tickets.remove(pt)

    if len(tickets) != 0 :
        writer.writerow([""])
        writer.writerow(["not in Sonny's"])
        for t in tickets : writer.writerow([str(t)])
