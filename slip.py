#! python3
# slip.py - retrieves pdf from adp and then scrapes data from that pdf

import time, pyautogui, os, tabula, csv, openpyxl, glob, pickle, logging, shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from datetime import timedelta

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')
logging.debug('Start of program')


if os.environ['COMPUTERNAME'] == 'LAPTOP':
    working_dir =   r'C:\Users\james\PycharmProjects\payroll_and_roster'
    slip_inbox =  ''
    gecko_driver = r"C:\Users\james\Google Drive\System Files\geckodriver.exe"

elif os.environ['COMPUTERNAME'] == 'JAMESPC':
    working_dir = r'A:\Python\payroll_and_roster'
    gecko_driver = r'A:\Program Files\geckodriver.exe'
    slip_inbox = r'C:\Users\james\Downloads'
else:
    Exception()


class Payslips:
    def __init__(self, password):
        self.url = r'https://secure.adppayroll.com.au/'
        self.work_ID = 'PR00002'
        self.employee_ID = '500435'
        self.password = password
        self.slips = []

    def get_payslip(self):
        browser = webdriver.Firefox(executable_path=gecko_driver)
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
        # put '1' between both halves to reviece most recent payslip.
        # put '2' to recieve second newest
        xpath_p1 = 'html[1]/body[1]/app-root[1]/div[1]/app-layout-component[1]/app-layout[1]/div[1]/div[2]/div[2]/div[2]/app-view-payslips[1]/div[1]/tile-content[1]/app-payslips-download[1]/adp-table[1]/table[1]/tbody[1]/tr['
        xpath_p2 = ']/td[3]/ng-component[1]/adp-button[1]'
        browser.find_element_by_xpath(xpath_p1 + '1' + xpath_p2).click()
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(2)
        browser.close()
        time.sleep(2)
        slip_file = most_recent_file(slip_inbox, 'pdf')
        new_slip_path = working_dir + '\\' + 'payslips'
        shutil.move(slip_file, new_slip_path)
        return new_slip_path

    def process_payslip(self, path):
        PAYSLIP = Payslip(path)
        PAYSLIP.process()
        self.slips.append(PAYSLIP)


class Payslip:
    def __init__(self, path):
        self.pdf_path = path
        self.scan_pdf()
        self.advice_table = self.get_advice_table()
        self.standard_table = self.get_standard_table()
        self.advice = {}

    def process(self):
        self.scan_pdf()
        self.extract_dates()
        self.compile_advice()
        self.compile_decuctions()
        self.compile_leave()
        self.get_hours_earned()
        self.get_hours_worked()
        self.remove_garbage()
        self.extract_pay()
        del self.advice_table
        del self.standard_table
        self.rename_slip()


    def scan_pdf(self):
        csv_filename = os.path.dirname(self.pdf_path) + '\\' + 'tempfile.csv'
        tabula.convert_into(self.pdf_path, csv_filename, output_format="csv", pages='all')
        file = open(csv_filename)
        tabReader = csv.reader(file)
        tablist = list(tabReader)
        self.raw_table =  tablist

    def extract_fortnight(self):
        return self.standard_table[2][0]

    def extract_dates(self):
        fornight_start = self.standard_table[2][0]
        fortnight_end = self.standard_table[2][1]
        format = '%d/%m/%Y'
        self.advice['begin'], self.advice['end'] = datetime.strptime(fornight_start, format), datetime.strptime(fortnight_end, format)
        return self.advice['begin'], self.advice['end']

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
        self.advice['advice'] = pay
        return pay

    def compile_decuctions(self):
        deductions = []
        for deduction in self.advice_table[:3]:
            deductions.append(deduction[-1])
        #super advice in first and third position
        self.advice['super'] = (deductions[0], deductions[2])
        self.advice['union_fee'] = deductions[1]
        return deductions

    def compile_leave(self):
        self.advice['annual_leave'] =    float(self.standard_table[12][5].split()[0])
        self.advice['long_service'] =    float(self.standard_table[14][4].split()[0])
        self.advice['EDO'] =             float(self.standard_table[16][5].split()[0])
        self.advice['holiday_credits'] = float(self.standard_table[15][5].split()[0])
        self.advice['SWOMC'] =           float(self.standard_table[19][4].split()[1])

        return self.advice

    def extract_pay(self):
        self.advice['net_pay'] = float(self.standard_table[9][5])
        self.advice['tax'] = float(self.standard_table[9][3].split()[-1])
        self.advice['gross_pay'] = float(self.standard_table[9][0])
        return self.advice['gross_pay'], self.advice['tax'], self.advice['net_pay']

    def get_hours_worked(self):
        target_cell = self.standard_table[19][2]
        # hrs earned is 6th word of target cell
        self.advice['hours_worked'] = float(target_cell.split()[5])
        return self.advice['hours_worked']

    def get_hours_earned(self):
        target_cell = self.standard_table[19][2]
        # hrs earned is 3rd word of target cell
        self.advice['hours_earned'] = float(target_cell.split()[2])
        return self.advice['hours_earned']

    def rename_slip(self):
        name = '\\Payslip '
        name = name + self.advice['end'].strftime("%d-%m-%Y") +'.pdf'
        name = os.path.dirname(self.pdf_path) + name
        print(name)
        os.rename(self.pdf_path, name)

    def remove_garbage(self):
        pass





class Roster:
    def __init__(self, roster_location):
        self.roster_location = roster_location
        if not hasattr(self, 'compiled_roster'):
            self.compiled_roster = self.roster_build()
        self.epoch = self.compiled_roster[0].epoch
        print(self.epoch)
        print(vars(self))


    def roster_build(self):
        compiled_roster = []
        if not hasattr(self, 'raw_roster'):
            self.raw_roster = self.get_raw_workbook()
        for day in range(14):
            x = RosterDay(self.raw_roster, day)
            x.extract_shifts()
            compiled_roster.append(x)
        del self.raw_roster
        return compiled_roster


    def get_raw_workbook(self):
        pickle_location = self.roster_location + '\\' + 'XL.pickle'
        if os.path.exists(pickle_location):
            print(pickle_location)
            with open(pickle_location, 'rb') as file:
                unpickled =  pickle.load(file)
                #TODO: add check to ensure the pickled object matches the newest XL spreadhseet
                return unpickled
        else:
            most_recent = most_recent_file(self.roster_location, 'xlsx')
            raw_roster = openpyxl.load_workbook(most_recent, data_only=True)
            with open(pickle_location, 'wb') as file:
                pickle.dump(raw_roster, file)
            return raw_roster





class RosterDay:
    def __init__(self, roster, day):
        self.raw_roster = roster
        self.date = self.raw_roster.worksheets[0]['F1'].value + timedelta(days=day)
        self.day_sheet = roster.worksheets[day]
        self.epoch = self.day_sheet['F1'].value
        #all dates must be calculated relative to the dat on sheet 0:
        self.shifts = []

    def extract_shifts(self):
        #find the last shift on the normal roster
        for i in range(5,self.day_sheet.max_row,4):
            cell = self.day_sheet['A' + str(i)]
            try:
                int(cell.value)
                last_cell = cell
            except:
                pass
        last_cell = self.get_last_cell()
        logging.debug('last cell is' + str(last_cell))


        jobs = []
        #cell is B5, then each new job is 4 row down
        #so we skip 4 rows each iteration
        for i in range(5,last_cell,4):
            job = {}
            job['id'] = self.get_cell_data(self.day_sheet, i, 'B') if self.get_cell_data(self.day_sheet, i, 'B') else None
            job['down'] = self.get_cell_data(self.day_sheet, i, 'C') if self.get_cell_data(self.day_sheet, i, 'C')  else None
            job['up'] = self.get_cell_data(self.day_sheet, i, 'E') if self.get_cell_data(self.day_sheet, i, 'E') else None
            job['dest'] = self.get_cell_data(self.day_sheet, i, 'D') if self.get_cell_data(self.day_sheet, i, 'D') else None
            job['start'] = int(self.get_cell_data(self.day_sheet, i, 'G')[0]) if self.get_cell_data(self.day_sheet, i, 'G') else None
            job['finish'] = int(self.get_cell_data(self.day_sheet, i, 'H')[0]) if self.get_cell_data(self.day_sheet, i, 'G') else None
            job['isrest'] = self.get_cell_data(self.day_sheet, i, 'J')[0] if self.get_cell_data(self.day_sheet, i, 'J') else False
            job['required'] = False if job['up'] and job['up'][0] in ['OFF', 'EDO'] else True
            jobs.append(job)
        self.shifts = jobs
        del self.day_sheet
        del self.raw_roster
        return self

    def get_last_cell(self):
        for i in range(5,self.day_sheet.max_row,4):
            cell = self.day_sheet['A' + str(i)]
            try:
                int(cell.value)
                last_cell = cell
            except:
                pass
        logging.debug('last cell is' + str(last_cell))
        return  last_cell.row + 1



    def name_positions(self):
        names = []
        last_cell = self.get_last_cell()
        for i in range(5,last_cell,4):
            names.append(self.get_cell_data(self.day_sheet, i, 'I')[0])
        return names



    def get_cell_data(self, sheet, offset, column):
        data = []
        # the cells is always 4 rows wide
        for line in range(offset,offset +4):
                cell = sheet[column+str(line)].value
                if cell is not None:
                    data.append(cell)
        return data



def most_recent_file(location, extention):
    #returns the most recent file with extention
    path = location + '\\' + extention
    list_of_files = glob.glob(path)  # * means all if need specific format then *.csv
    print(list_of_files)
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file


#password = input('input passy')

def most_recent_file(location, extention):
    #returns the most recent file with extention
    path = location + '\\*.' + extention
    list_of_files = glob.glob(path)  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

if __name__ == '__main__':
    if os.path.exists(working_dir+'\\'+'roster.pickle'):
        with open(working_dir+'\\'+'roster.pickle', 'rb') as file:
            roster = pickle.load(file)
    else:
        roster = Roster(working_dir)
    print(roster.compiled_roster[0].shifts)
    payslips = Payslips('t8C3&Lq9uTy0')
    #payslips.get_payslip()
    slip = most_recent_file(working_dir + '\\payslips', 'pdf')
    print(slip)
    payslips.process_payslip(slip)

    with open(working_dir+'\\'+'roster.pickle', 'wb') as file:
            pickle.dump(roster, file)


