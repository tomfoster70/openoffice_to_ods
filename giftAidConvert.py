from collections import namedtuple
import csv
from datetime import *
from ezodf import opendoc
import os
import sys
import shutil
import re
import time
import traceback
# For xlt reading
import xlrd as xlrd

transactionLine = namedtuple('transactionLine', ['date', 'firstName', 'surname', 'amount'])


class ExitException(Exception):
    pass


def output_to_ods(all_output, firstDate):
    doc = opendoc("Forms/gift_aid_schedule.ods")

    doc.docname = "output/gift_aid_schedule_output.ods"
    sheet = doc.sheets[0]

    print(firstDate)
    do = datetime.strptime("01/01/01", "%d/%m/%y")

    delta = (firstDate - do).days + 367
    sheet[12, 3].set_value(delta)

    with open('input/unclaimed.csv', 'w') as unclaimedNew:
        csv_unclaimed_writer = csv.writer(unclaimedNew, dialect='excel')

        i = 0
        for output in all_output:
            if output.address == "":
                csv_row = [ output.lastDate, output.firstName, output.surname, round(output.amount, 2) ]
                csv_unclaimed_writer.writerow(csv_row)
            else:
                sheet[i + 24, 2].set_value(output.title)
                sheet[i + 24, 3].set_value(output.firstName)
                sheet[i + 24, 4].set_value(output.surname)
                sheet[i + 24, 5].set_value(output.address)
                sheet[i + 24, 6].set_value(output.postcode)

                dt = datetime.strptime(output.lastDate, "%d/%m/%Y")
                delta = (dt - do).days + 367
                sheet[i + 24, 9].set_value(delta)
                sheet[i + 24, 10].set_value(round(output.amount, 2))
                i += 1

    doc.save()


def output_to_csv(all_output, firstDate, totalFromReports, outputFilename='output/OutputGiftAidSpreadsheet.csv'):
    with open(outputFilename, 'wb') as csvfileNew:
        csv_writer = csv.writer(csvfileNew, dialect='excel')

        with open('input/unclaimed.csv', 'wb') as unclaimedNew:
            csv_unclaimed_writer = csv.writer(unclaimedNew, dialect='excel')

            csv_writer.writerow(["FirstDate: ", firstDate, "Total from reports", totalFromReports, "Gift Aid Amount: ", totalFromReports * 0.25])

            for output in all_output:

                if output.address == "":
                    csv_row = [ output.lastDate, output.firstName, output.surname, output.amount ]
                    csv_unclaimed_writer.writerow(csv_row)
                else:
                    csv_row = [ output.title, output.firstName, output.surname, output.address, output.postcode, "", "", output.lastDate, output.amount ]
                    csv_writer.writerow(csv_row)


class outputLine():

    def __init__(self, date, title, firstname, surname, amount, address, postcode):
        self.lastDate = date
        self.title = title
        self.firstName = firstname
        self.surname = surname
        self.amount = float(amount)
        self.address = address
        self.postcode = postcode


class GiftAidReport():

    def __init__(self, report):
        self.report = report

        # Headers
        self.date_offset = None
        self.deposit_offset = None
        self.name_offset_start = None
        self.name_offset_fin = None
        self.total_cell_offset = None
        self.total_gift_aid_amount_offset = None
        self.total_balance_offset = None

        self.addressbook = "AddressBook.csv"

        # Other items
        self.totalFromTrans = 0
        self.totalFromReport = 0
        self.transactions = []
        self.outputLines = []

        if "unclaimed.csv" in report:
            self.unclaimed = True
        else:
            self.unclaimed = False
            with xlrd.open_workbook(self.report, 'rb') as fh:
                self.sheet = fh.sheet_by_index(0)

        self.process_report()

    def get_headers(self):
        print("Looking at {}".format(self.report))

        found_headers = False
        for row in range(self.sheet.nrows):
            try:
                for i in range(0, 29, 1):
                    value = self.sheet.cell(row, i).value
                    if 'date' in value.lower() and self.date_offset == None:
                        self.found_headers = True
                        self.date_offset = i
                    if 'type' in value.lower()  and self.deposit_offset == None:
                        self.deposit_offset = i
                    if 'name' in value.lower()  and self.name_offset_start == None:
                        self.name_offset_start = i
                    if 'split' in value.lower()  and self.name_offset_fin == None:
                        self.name_offset_fin = i
                    if 'amount' in value.lower()  and self.total_cell_offset == None:
                        self.total_cell_offset = i
                    if 'total' in value.lower() and 'gift aid income' in value.lower() and self.total_gift_aid_amount_offset == None:
                        self.total_gift_aid_amount_offset = i
                    if 'balance' in value.lower()  and self.total_balance_offset == None:
                        self.total_balance_offset = i
            except Exception as e:
                # print("Exception {}".format(e))
                pass

        if self.found_headers is False:
            raise ExitException("Failed to get headers")

        print('header offsets')
        print(self.date_offset, ' Date ')
        print(self.deposit_offset, ' Deposit/type ')
        print(self.name_offset_start, ' name start ')
        print(self.name_offset_fin, ' name fin ')
        print(self.total_cell_offset, ' total/amout ')
        print(self.total_balance_offset, ' balance ')
        print(self.total_gift_aid_amount_offset, ' where the total name is')

    def parse_name(self, name):
        nameItems = name.split(' ')
        firstName = nameItems[0]
        surname = " ".join(nameItems[1:])
        surname = surname.split(' ')[-1]

        return firstName, surname

    def get_name(self, row):
        payeename = []

        name = self.sheet.cell(row, self.name_offset_start).value

        if name != "":
            firstName, surname = self.parse_name(name)
        else:
            firstName = ""
            surname = ""
            # Try getting the name from the description
            for i in range(self.name_offset_start + 1, self.name_offset_fin, 1):
                temp = self.sheet.cell(row, i).value.split(',')
                for item2 in temp:
                    payeename.append(item2)
            name = " ".join(payeename)

            print("Name before remove: {}".format(name))
            items_to_remove = ['chq in offering', '\d+/\d+/\d+', 'Church', 'Collection', 'Deposit', 'New',
                'Godfirst', 'Special', 'chq', 'Chq', 'Stewardship', 'Services',
                'online', 'giving', 'cash', 'Cheque', 'Izettle', 'For Barry Church', ':', '[F|f]or .*']
            for item in items_to_remove:
                name = re.sub(item, '', name)
            # name = re.sub('chq in offering', '', name)
            # name = re.sub('\d+/\d+/\d+', '', name)
            name = name.lstrip()
            name = name.rstrip()

            print("NAME after remove", name)
            firstName, surname = self.parse_name(name)
            print(firstName, surname)

        return firstName, surname

    def process_giving_xlsx(self):

        get_contents = 1
        for row in range(self.sheet.nrows):

            deposit = self.sheet.cell(row, self.deposit_offset).value
            date_cell = self.sheet.cell(row, self.date_offset).value
            amount_payee_cell = self.sheet.cell(row, self.total_cell_offset).value
            total_cell = self.sheet.cell(row, self.total_cell_offset).value

            if 'total' in self.sheet.cell(row, self.total_gift_aid_amount_offset).value.lower() \
                and 'gift aid income' in self.sheet.cell(row, self.total_gift_aid_amount_offset).value.lower() \
                and 'non' not in self.sheet.cell(row, self.total_gift_aid_amount_offset).value.lower():

                print("row", row, total_cell, self.sheet.cell(row, self.total_balance_offset).value)
                self.totalFromReport += total_cell
                get_contents = 0
                continue
            elif deposit != 'Deposit':
                continue
            elif get_contents == 1 and deposit == 'Deposit':
                try:
                    date = xlrd.xldate_as_tuple(date_cell, fh.datemode)
                    date = "%s/%s/%s" % (date[2], date[1], date[0])
                except:
                    date = date_cell

                # need to get amaount, also remove ',' from amounts over 1000.
                amount = str(amount_payee_cell)

                firstName, surname = self.get_name(row)

                self.totalFromTrans += float(amount)
                if firstName == "":
                    raise ExitException("firstName is NULL {} {}".format(row,))
                self.transactions.append(transactionLine(date, firstName, surname, amount))

    def process_giving_csv(self):

        with open(self.report) as fh:
            self.unclaimed_csv_reader = csv.reader (fh)
            for row in self.unclaimed_csv_reader:
                try:
                    date = row[0].strftime("%m/%d/%Y")
                    firstName = row[1]
                    surname = row[2]
                    amount = row[3]
                    self.transactions.append(transactionLine(date, firstName, surname, amount))
                except:
                    continue

    def get_propername_and_address(self, trans, transFirstName, transSurname):

        found = False
        with open(self.addressbook) as fh:
            self.addressbook_csv_reader = csv.reader (fh)
            for row in self.addressbook_csv_reader:

                title = row[0]
                firstname = row[1]
                surname = row[2]
                address = row[3]
                postcode = row[4]

                if surname.lower() == transSurname.lower():

                    if '&' in transFirstName:
                        firstNames = transFirstName.split('&')
                    else:
                        firstNames = [transFirstName]

                    for fn in firstNames:
                        if fn == firstname:
                            return outputLine(trans.date, title, firstname, surname, trans.amount, address, postcode)
                        elif fn[0].lower() == firstname[0].lower():
                            return outputLine(trans.date, title, firstname, surname, trans.amount, address, postcode)

        return None

    def process_report(self):

        if self.unclaimed == True:
            self.process_giving_csv()
        else:
            self.get_headers()
            self.process_giving_xlsx()

        print(self.transactions)
        for trans in self.transactions:
            # Normal
            ol = self.get_propername_and_address(trans, trans.firstName, trans.surname)

            # Surname is the first name
            if ol is None:
                ol = self.get_propername_and_address(trans, trans.surname, trans.firstName)

            # empty line
            if ol is None:
                ol = outputLine(trans.date, "", trans.firstName, trans.surname, trans.amount, "", "")

            self.outputLines.append(ol)

        if round(self.totalFromReport, 2) != round(self.totalFromTrans, 2):
            raise ExitException("Gift aid total from {} don't match what I calculated they should be {} != {}".format(self.report, self.totalFromReport, self.totalFromTrans))


def first_date_calc(firstDate, lastDate):

    outputDate = datetime.strptime(lastDate, "%d/%m/%Y")
    if firstDate is None:
        firstDate = outputDate
    else:
        if outputDate < firstDate:
            firstDate = outputDate

    return firstDate


def main():
    reports = []
    for item in os.listdir("input"):
        print("looking at {}".format(item))
        file_path = os.path.join("input", item)
        claim_path = os.path.join("claimed", item)

        # Get report info
        reports.append(GiftAidReport(file_path))

        # Move to claimed
        if "unclaimed.csv" not in item:
            shutil.move(file_path, claim_path)

    if reports == []:
        raise ExitException("Nothing to process - no files at all in the input folder?")

    all_output = []
    firstDate = None
    totalFromReports = 0

    # Loop all of the reports and count everything up
    for report in reports:
        totalFromReports += report.totalFromReport

        # For line in report
        for output in report.outputLines:
            print(type(output.lastDate), output.lastDate, output.firstName, " :sur: ", output.surname, output.amount, output.address, output.postcode)

            found = False

            # Check to see if we have already got this person in the output
            for r in all_output:
                if output.firstName == r.firstName and output.surname == r.surname:
                    outputDate = datetime.strptime(output.lastDate, "%d/%m/%Y")
                    rDate = datetime.strptime(r.lastDate, "%d/%m/%Y")

                    if outputDate < rDate:
                        r.lastDate = output.lastDate

                    firstDate = first_date_calc(firstDate, output.lastDate)

                    r.amount = r.amount + output.amount
                    found = True
                    break

            # If we haven't seen this person before then add them!
            if found is False:

                firstDate = first_date_calc(firstDate, output.lastDate)

                all_output.append(output)

    print("----------")
    if all_output == []:
        raise ExitException("Nothing to output - are there any input files?")

    # output_to_csv(all_output, firstDate, totalFromReports)
    output_to_ods(all_output, firstDate)
    for output in all_output:
        print(output.lastDate, output.firstName, " :sur: ", output.surname, output.amount, output.address, output.postcode)


if __name__ == "__main__":
    print('Starting...')
    try:
        main()
    except Exception as e:
        print('Failed for some reason - please ask tom to fix me. Please include input files and this error [{}]'.format(e))
        exc_type, exc_value, exc_traceback = sys.exc_info()
        traceback.print_tb(exc_traceback)
        time.sleep(20)
