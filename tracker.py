import xlrd
import xlwt
from datetime import datetime
from datetime import timedelta
from time import sleep
import win32com.client
from classes import *

def date_convert(excelDate) :
    # if excelDate == "N/A" or excelDate == "" : return None
    seconds = (excelDate - 25569) * 86400.0
    return datetime.utcfromtimestamp(seconds)

'''
Read data from iipm or iipm extract and return dictionary

The data structure (app_bin) is set as dictionary `string` : `list of class`
    app_code : [App, Contact (SLA), Contact (Meeting)]

'''
def setup(iipm, status, contact, meeting) :
    app_bin = {}

    # Find apps we want to track from IIPM
    for i in range(1, iipm.nrows) :
        phase = iipm.row(i)[8].value
        if phase != 'Retired' :
            c = iipm.row(i)[0].value
            # Take care of number-only app code
            if isinstance(c, float) or isinstance(c, int) : c = str(int(c))
            if c == "" :
                print("No Appcode for row " + str(i))
            elif c not in app_bin :
                an = iipm.row(i)[3].value
                custodian = iipm.row(i)[4].value
                l5 = iipm.row(i)[5].value

                app_bin[c] = [App(c, an, custodian, phase)]
            else :
                print("Duplicate data : " + c)
    print('end')

    # SLA Overall Status Info
    for i in range(1, status.nrows) :
        c = status.row(i)[0].value
        # Take care of number-only app code
        if isinstance(c, float) or isinstance(c, int) : c = str(int(c))

        st = status.row(i)[7].value
        start = ""
        end = ""
        fr = ""

        if c in app_bin:
            if st == 'N/A' or st == 'In Progress' or st == 'Not Started':
                app_bin[c][0].set_status(st)
            elif st == "" and status.row(i)[3].value == "":
                app_bin[c][0].set_status('Not Started')
            else :
                start = status.row(i)[3].value
                end = status.row(i)[4].value
                fr = status.row(i)[5].value

                # Excel date is in number format so change it to string
                if isinstance(start, float) :
                    start = date_convert(start).strftime("%Y/%m/%d")
                if isinstance(end, float) :
                    end = date_convert(end).strftime("%Y/%m/%d")

                if start != "" and end != "" :
                    if datetime.today() > datetime.strptime(end, "%Y/%m/%d") :
                        app_bin[c][0].set_status("Expired")
                    elif datetime.today() + timedelta(days=30) > \
                    datetime.strptime(end, "%Y/%m/%d") :
                        app_bin[c][0].set_status("Expiring within a month")
                    else :
                        app_bin[c][0].set_status("In Good Standing")

            app_bin[c][0].add_slaInfo(start, end, fr)

        else :
            print(c + " not in our list.  Check Again")

    # Expiration Contact Info
    for i in range(1, contact.nrows) :
        c = contact.row(i)[0].value
        if isinstance(c, float) or isinstance(c, int) : c = str(int(c))
        ctd = contact.row(i)[1].value
        rpd = contact.row(i)[2].value
        td = contact.row(i)[3].value
        done = contact.row(i)[4].value
        lcd = contact.row(i)[5].value

        # Excel date is in number format so change it to string
        if isinstance(ctd, float) :
            ctd = date_convert(ctd).strftime("%Y/%m/%d")
        if isinstance(rpd, float) :
            rpd = date_convert(rpd).strftime("%Y/%m/%d")
        if isinstance(td, float) :
            td = date_convert(td).strftime("%Y/%m/%d")
        if isinstance(lcd, float) :
            lcd = date_convert(lcd).strftime("%Y/%m/%d")

        if c in app_bin :
            app_bin[c].append(Contact(c, ctd, lcd, rpd, td))
            if td == "" :
                app_bin[c][1].set_status("Not Contacted")
            elif datetime.today() > datetime.strptime(td, "%Y/%m/%d") :
                app_bin[c][1].set_status("Overdue")
            elif datetime.today() <= datetime.strptime(td, "%Y/%m/%d") :
                app_bin[c][1].set_status("Contacted")
        #TODO New app detected so write in file?
        #else :

    return app_bin


def find_who_to_contact(app_bin) :
    custodian = {}

    for app in app_bin :
        if app_bin[app][0].status == "Expired" or \
        app_bin[app][0].status == "Expiring within a month" and \
        app_bin[app][1].status != "Contacted":
            if app_bin[app][0].appCustodian in custodian :
                custodian[app_bin[app][0].appCustodian].append(app_bin[app][0])
            else :
                custodian[app_bin[app][0].appCustodian] = [app_bin[app][0]]

    return custodian

# Only for one app custodian
def construct_email(apps) :
    app_str = ""
    expired = False
    acFirstName = apps[0].appCustodian.split(',')[1]

    for app in apps :
        app_str += app.code + ", "
        if app.status == "Expired" : expired = True

    due = 3 if expired else 7
    target = datetime.today() + timedelta(days=due)

    body = "Hello, " + acFirstName + \
        "\n\nI am reaching out to you for soon to be / expired SLAs " + \
        "for following apps: " + app_str[:-2] + "\n\n" + \
        "A kind reminder to include measureable KPI’s if possible, " + \
        "service quality review meeting frequency (monthly or quarterly) " + \
        "and its dates are a must.\n\n" + \
        "Also SLA duration does not have to be one year." + \
        "For your convenience, it can be 1 year and 2 months, etc.\n" + \
        "The purpose is that you can spread out expiring dates.\n\n" + \
        "We have blank template, if you need one we can provide it.\n\n" + \
        "If you can provide us an update and target completion date, " + \
        "that would be much appreciated.\n\n" + \
        "Please reply by " + target.strftime("%Y/%m/%d") + "\n\n" + \
        "Thank you" + \
        "\n\nThis is sent from Python\n\n"

    return (app_str[:-2], body)

def sendEmail(to, cc, subject, body) :
    outlook = win32com.client.Dispatch("outlook.application")

    mail = outlook.CreateItem(0)
    mail.To = to
    mail.CC = cc
    mail.Subject = subject
    mail.body = body
    mail.send

def readLastEmail() :
    outlook = win32com.client.Dispatch("outlook.application").getNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items
    message = messages.GetLast().Get
    print(message.body)

if __name__ == "__main__" :

    workbook = xlrd.open_workbook('test.xlsx')

    sheetIIPM = workbook.sheet_by_name('IIPM')
    sheetStatus = workbook.sheet_by_name('SLAStatus')
    sheetContact = workbook.sheet_by_name('ContactLog')
    sheetMeeting = workbook.sheet_by_name('Meeting')

    app_bin = setup(sheetIIPM, sheetStatus, sheetContact, sheetMeeting)
    #for app in app_bin : print(app + ' ' +app_bin[app][0].slaEnd + ' ' + app_bin[app][0].status)
    contact_list = find_who_to_contact(app_bin)
    i = 0
    for custodian in contact_list :
        '''
        app_str = ""
        for app in contact_list[custodian] : app_str += app.code + ", "
        print(custodian + ' - ' + app_str[:-2])
        '''
        title = "SLA Expiry Notice - "
        apps, body = construct_email(contact_list[custodian])
        title += apps

        to = "elisa.yamchou@rbc.com"
        cc = ""
        b = "Delete this - it's from Python :)\n\n" + body
        #cc = "carl.morgan@rbc.com; dragan.jeremic@rbc.com"
        #sendEmail(to, cc, title, b)
        i += 1
        print("Email sent to " + custodian)
        # Pause 1 second so you don't spam people
        sleep(1)
