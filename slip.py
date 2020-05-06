#! python3
# slip.py - retrieves pdf from adp and then scrapes data from that pdf

import time, pyautogui, os, tabula, csv, openpyxl, glob, pickle
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime


if os.environ['COMPUTERNAME'] == 'LAPTOP':
    working_dir =   r'C:\Users\james\Downloads'
    slip_folder =       r'C:\Users\james\PycharmProjects\payroll_and_roster'

elif os.environ['COMPUTERNAME'] == 'JAMESPC':
    working_dir = r'A:\Python\payroll_and_roster'
    download_folder = ''
else:
    Exception()


class payslip:
    def __init__(self, path):
        self.pdf_path = path = "\\" + 'payslips'
        self.scan_pdf()
        self.advice_table = self.get_advice_table()
        self.standard_table = self.get_standard_table()

    def scan_pdf(self):
        csv_filename = os.path.dirname(self.pdf_path) + '\\' + 'tempfile.csv'
        print(csv_filename)
        tabula.convert_into(os.path.basename(self.pdf_path), csv_filename, output_format="csv", pages='all')
        file = open("tempfile.csv")
        tabReader = csv.reader(file)
        tablist = list(tabReader)
        self.raw_table =  tablist

    def extract_dates(self):
        fornight_start = self.table[2][0]
        fortnight_end = self.table[2][1]
        format = '%d/%m/%Y'
        return datetime.strptime(fornight_start, format), datetime.strptime(fortnight_end, format)

    def find_payment_table(self):
        #returns the range the pay advice table occupies
        for index, line in enumerate(self.raw_table):
            if "Description" in line:
                advice_begins = index + 1
                break
        for index, line in enumerate(self.raw_table):
            if "SUMMARY OF EARNINGS" in line:
                advice_ends = index
                break

        return advice_begins, advice_ends


    def get_advice_table(self):
        begins, ends = self.find_payment_table()
        return self.raw_table[begins:ends]

    def get_standard_table(self):
        begins, ends = self.find_payment_table()
        table = self.raw_table[:begins] + self.raw_table[ends:]
        return table

    def compile_advice(self):
        pay = []
        for pay_type in self.advice_table:
            fixedCell4 = pay_type[3].split()[0]
            data = pay_type[:3] + [fixedCell4]
            for i,x in enumerate(data[1:]):
                data[i+1] = float(x)
            pay.append(data)
        return pay

    def compile_decuctions(self):
        deductions = []
        for deduction in self.advice_table[:2]:
            duction = deduction[3].split()[1:]
            deductions.append((' '.join(duction),float(deduction[-1])))
        return deductions

    def compile_leave(self):
        target_rows = ('LSL Full', 'PH Credits', 'EDO', 'Sick Full')
        leave = dict()
        for line in self.standard_table:
            for type in target_rows:
                if type in line:
                    for cell in line:
                        pass


    def get_hours_worked(self):
        pass

    def get_hours_earned(self):
        pass



class Roster:
    def __init__(self, roster_location):
        self.roster_location = roster_location
        if not hasattr(self, 'compiled_roster'):
            self.compiled_roster = self.roster_build()

    def roster_build(self):
        if not hasattr(self, 'raw_roster'):
            pickle_location = self.roster_location + '\\' + 'XL.pickle'
            if os.path.exists(pickle_location):
                with open(pickle_location, 'rb') as file:
                    self.raw_roster = pickle.load(file)
            else:
                most_recent = most_recent_file(self.roster_location, 'xlsx')
                self.raw_roster = openpyxl.load_workbook(most_recent)
                with open(pickle_location, 'wb') as file:
                    pickle.dump(self.raw_roster, file)

        print(self.raw_roster)
        for day in range(14):
            pass
            #fortnight.append(self.build_day(roster.get))
        #return fortnight

    def load_XL_book(self):
        most_recent = most_recent_file(self.roster_location, 'xlsx')
        roster = openpyxl.load_workbook(most_recent)
        return roster









class Payslips:
    def __init__(self, password):
        self.url = r'https://secure.adppayroll.com.au/'
        self.work_ID = 'PR00002'
        self.employee_ID = '500435'
        self.password = password
        now = time.time() - 60*60*24*14



    def get_payslip(self):
        browser = webdriver.Firefox(executable_path=r"C:\Users\james\Google Drive\System Files\geckodriver.exe")
        browser.get(self.url)
        workID_elem = browser.find_element_by_id('client_id')
        employeeID_elem = browser.find_element_by_id('user_id')
        password_elem = browser.find_element_by_id('password')
        employeeID_elem.send_keys(self.employee_ID)
        workID_elem.send_keys(self.work_ID)
        password_elem.send_keys(self.password)
        password_elem.send_keys(Keys.ENTER)
        time.sleep(4)
        browser.get(r'https://my.adppayroll.com.au/webapp/paydetails/payslips')
        time.sleep(8)
        #put '1' between both halves to reviece most recent payslip.
        #put '2' to recieve second newest
        xpath_p1 = 'html[1]/body[1]/app-root[1]/div[1]/app-layout-component[1]/app-layout[1]/div[1]/div[2]/div[2]/div[2]/app-view-payslips[1]/div[1]/tile-content[1]/app-payslips-download[1]/adp-table[1]/table[1]/tbody[1]/tr['
        xpath_p2 = ']/td[3]/ng-component[1]/adp-button[1]'
        browser.find_element_by_xpath(xpath_p1 + '1' + xpath_p2).click()
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(2)
        browser.close()


def most_recent_file(location, extention):
    #returns the most recent file with extention
    path = location + '/' + extention
    list_of_files = glob.glob(path)  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file


#password = input('input passy')
r'''
p = payslip(r'C:\Users\james\PycharmProjects\payroll_and_roster\Q001830A.2020429.95147.0001453249.pdf')
#p.extract_dates()
[print(x) for x in p.standard_table]
print('#####################')
[print(x) for x in p.compile_decuctions()]


def load():
    if 'notable_helper.pickle' in os.listdir():
        with open('notable_helper.pickle', 'rb') as file:
            manager = pickle.load(file)

    else:
        manager = System(original_folder, modified_folder)


def save():
    with open('notable_helper.pickle', 'wb') as file:
        pickle.dump(manager, file)

'''

def most_recent_file(location, extention):
    #returns the most recent file with extention
    path = location + '\\*.' + extention
    print(path)
    list_of_files = glob.glob(path)  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

roster = Roster(working_dir)