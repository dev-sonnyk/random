import xlrd
import xlwt

class Task :

    def __init__(self, id, title, open_time, status, parent) :
        self.id = id
        self.title = title
        self.ot = open_time
        self.status = status
        self.parent = parent

read = xlrd.open_workbook('Copy of IM PM and Severity Chart.xlsx')
write = xlwt.Workbook(encoding = "ascii")
new_sheet = write.add_sheet('PMTasksJoined')

workbook = read.sheet_by_name('PM Tasks Table')

pm = {}

for i in range(1, workbook.nrows) :
    task = workbook.row(i)[1].value
    title = workbook.row(i)[2].value
    ot = workbook.row(i)[3].value
    print(ot)
    st = workbook.row(i)[4].value
    pr = workbook.row(i)[5].value
    if task != "" :
        t = Task(task, title, ot, st, pr)

        if pr not in pm :
            pm[pr] = t
        else :
            pm[pr].id += ",\n"
            pm[pr].id += task
            pm[pr].title += ",\n"
            pm[pr].title += title
            pm[pr].ot += ",\n"
            pm[pr].ot += ot
            pm[pr].status += ",\n"
            pm[pr].status += st

i = 0
for pr in pm :
    new_sheet.write(i, 0, label = pr)
    new_sheet.write(i, 1, label = pm[pr].id)
    new_sheet.write(i, 2, label = pm[pr].title)
    new_sheet.write(i, 3, label = pm[pr].ot)
    new_sheet.write(i, 4, label = pm[pr].status)
    i += 1

# It will not overwrite so delete the old file or put new name
write.save('Concatenated2.xls')
