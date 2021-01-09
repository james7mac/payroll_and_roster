#! python3
# slip.py - retrieves pdf from adp and then scrapes data from that pdf
#TODO: when updating roster you must input the number of full pages below as well as the coords for
# final page
# The roster must also be cleaned of all blue "available" this was done using 'Libre Draw'
pdfFullPages = 9
pdfLastCoords = [148.6, 107.0 ,1166.4, 251.5]
epoch = '30/06/21' #date in 30/6/99 format
import time, os, tabula, csv, openpyxl, glob, pickle, logging, shutil, holidays
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from datetime import timedelta
from datetime import time as TIME
from googlecal import update_calander, check_work_event, delete_event, get_creds

epoch = datetime.strptime(epoch,'%d/%m/%y')
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')
logging.debug('Start of program')

public_hols = holidays.CountryHoliday('Australia', prov="VIC")
for d,n in sorted(holidays.CountryHoliday('Australia', prov="VIC", years=2020).items()):
    print('{}   {}'.format(d,n))
if os.environ['COMPUTERNAME'] == 'JMLAPTOP':
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
        self.extract_pay()
        self.rename_slip()
        self.remove_garbage()

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
        self.advice['ending'], self.advice['pay date'] = datetime.strptime(fornight_start, format), datetime.strptime(fortnight_end, format)
        return self.advice['ending'], self.advice['pay date']

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
        name = name + self.advice['ending'].strftime("%d-%m-%Y") +'.pdf'
        name = os.path.dirname(self.pdf_path) + name
        print(name)
        os.rename(self.pdf_path, name)

    def remove_garbage(self):
        del self.advice_table
        del self.standard_table



class pRoster:
    def __init__(self, roster_location, compiled_roster=None):
        self.roster_location = roster_location
        if not compiled_roster:
            roster_pdf = open(self.roster_location, 'rb')
            self.compiled_roster = self.build_roster_dummy(roster_pdf)

    def build_roster(self, roster_pdf):
        tabula_roster = tabula.read_pdf(roster_pdf, pages='all')

    def build_roster_dummy(self, roster_pdf):
        if os.path.exists(working_dir+'\\TEMPFILE.csv'):
            with open(working_dir + "\\TEMPFILE.csv") as file:
               self.tabula_roster = csv.reader(file)


        else:
            tabula.convert_into(roster_pdf, 'TEMPFILE.csv', output_format="csv", pages='all')
            with open(working_dir + "\\TEMPFILE.csv") as file:
                self.tabula_roster = csv.reader(file)


    ## PANDAS DATA FRAME ??
class Roster:
    def __init__(self, roster_location, pdfFullPages, pdfLastCoords, epoch):
        self.roster_location = roster_location
        self.fix_cell_helper = 0
        self.pdfFullPages = pdfFullPages
        self.epoch = epoch
        self. pdfLastCoords = pdfLastCoords
        if not hasattr(self, 'dfroster'):
            self.df = self.roster_build(self.roster_location, pdfFullPages, pdfLastCoords)




    def roster_build(self, pdfPath, pdfFullPages, pdfLastCoords):
        df1 = pd.DataFrame()
        for i in range(1, pdfFullPages+1):
            roster_pdf = open(pdfPath, 'rb')
            # coords of the area in the roster found using GIMP:

            #if i == pdfFullPages+1:
            #    x = tabula.read_pdf(roster_pdf, pages=int(i), lattice=1, area=pdfLastCoords)
            #else:
            x = tabula.read_pdf(roster_pdf, pages=int(i), lattice=1, area=[100, 149, 589, 1167.9])
            if type(x) == list:
                x = x[0]
            x.dropna(axis=0, thresh=30, inplace=True)
            x.dropna(axis=1, how='all', inplace=True)
            x.columns = [str(i) for i in range(1, 15)] + ['h1'] + [str(i) for i in range(15, 29)] + ['h2']
            df1 = df1.append(x)
        df1.index = [str(i + 1) for i in range(len(df1))]
        df1.drop(['h1', 'h2'], axis=1, inplace=True)
        df1 = df1.applymap(self.fix_cell)
        df = pd.DataFrame()
        for i in df1.iterrows():
            for ii, dic in enumerate(i[1]):
                dic['rosDay'] = ii + 1
                dic['rosLine'] = int(i[0])
                temp_df = pd.DataFrame(dic, index=[0])
                df = df.append(temp_df)

        return df

    def fix_times(self, timeStr):
        if timeStr == '2400':
            timeStr = '0000'
        return datetime.strptime(timeStr, "%H%M").time()

    def fix_cell(self, x):
        self.fix_cell_helper += 1
        try:
            if x.startswith('OFF'):
                items = {'id': 'OFF', 'hours': None, 'start': None, 'finish': None, 'rest': False}
            elif x.startswith('SPE'):
                items = x.split('\r')
                fixTimes = items[2].split('-')
                del items[2]
                fixTimes = [self.fix_times(i) for i in fixTimes]
                items = {'id': items[0], 'hours': items[1], 'start': fixTimes[0], 'finish': fixTimes[1], 'rest': False}
                if x.startswith('SPEX'):
                    items['rest'] = True
            elif x.startswith('EDO'):
                items = {'id': 'EDO', 'hours': None, 'start': None, 'finish': None, 'rest': False}
            elif x.lower().startswith('av'):
                start, finish = x.split('\r')[-1].split('-')
                job = {'id': 'AV', 'hours': '8:00', 'start': self.fix_times(start), 'finish': self.fix_times(finish),
                       'rest': False}
                return job
            else:
                print([x])
                return x
            return items
        except(IndexError):
            print('INDEX ERROR PASSING FOLLOWING:  ' + x)
            print('col ' + str(self.fix_cell_helper // len(self.df)))
            print('row ' + str(self.fix_cell_helper % len(self.df)))
            return x



    def get_raw_workbook(self):
        pickle_location = self.roster_location + '\\' + 'XL.pickle'
        if os.path.exists(pickle_location):
            print(pickle_location)
            with open(pickle_location, 'rb') as file:
                unpickled =  pickle.load(file)
                #TODO: add check to ensure the pickled object matches the newest XL spreadhseet
                return unpickled
        else:
            print(self.roster_location)
            most_recent = most_recent_file(self.roster_location, 'xlsx')
            raw_roster = openpyxl.load_workbook(most_recent, data_only=True)
            with open(pickle_location, 'wb') as file:
                pickle.dump(raw_roster, file)
            return raw_roster

    def generate_roster(self, name, weeks_ahead):
        position = RosterDay.name_list.index(name)
        since_epoch = datetime.now() - RosterDay.epoch
        #times by 7 for weeks in days
        forcast_range = (7*weeks_ahead)
        indiv_roster = []
        for day in range(forcast_range+since_epoch.days):
            target_shift = position + (day//28)
            target_day = day%28
            shift = self.compiled_roster[target_day].shifts[target_shift]
            shift['date'] = RosterDay.epoch + timedelta(days=day)
            indiv_roster.append(shift)
        return indiv_roster

    def create_calander_event(self, job, service):
        existing_event = check_work_event(job['date'], service)
        print(self.format_job(job))
        if existing_event:
            delete_event(existing_event, service)
        update_calander(self.format_job(job), service)
        logging.debug('CAL UPDATE: ' + str(job['date']))

    def update_calander(self, jobs, service):
        print(jobs)
        for job in jobs:
            if job['required']:
                self.create_calander_event(job, service)
            else:
                existing_event = check_work_event(job['date'], service)
                if existing_event:
                    delete_event(existing_event, service)

    def add_destination(self, job, destination):
        offset = 0
        job2 = ''
        for i, j in enumerate(job):
            if 'prep' in j.lower():
                offset += 1
                continue
            else:
                try:
                    dest = destination[i - offset] if type(destination) == list else destination
                    job2 += str(job[i]) + ',  dest: ' + str(dest) + '\n'
                except:
                    job2 += "\ndestError, Raw destination info:\n" + str(destination)
        return job2

    def format_job(self, job):
        if not job['required']:
            return None
        formatted = {}
        if not job['isrest']:
            formatted['summary'] = "Work"
        else:
            formatted['summary'] = job['isrest']
        if type(job['down']) == list:
            down = self.add_destination(job['down'], job['down dest'])
        else:
            down = str(job['down']) + '   dest: ' + str(job['down dest'])
        if type(job['up']) == list and type(job['up dest']) == list:
            up = self.add_destination(job['up'], job['up dest'])
        else:
            up = str(job['up']) + '   dest: ' + str(job['up dest'])

        formatted['description'] = "DOWN\n{0}\n\n UP\n{1}".format(
            down, up)
        start_string = str(job['start']).rjust(4, '0')
        end_string =  str(job['finish']).rjust(4, '0')
        start = datetime.combine(job['date'], TIME(int(start_string[:2]), int(start_string[2:])))
        end = datetime.combine(job['date'], TIME(int(end_string[:2]), int(end_string[3:])))
        if end.hour in [0,1]:
            end = end + timedelta(days=1)
        formatted['start'] = {'dateTime': start.isoformat(), 'timeZone':'Australia/Melbourne'}
        formatted['end'] = {'dateTime': end.isoformat(), 'timeZone':'Australia/Melbourne'}
        formatted['description'] = formatted ['description'] + '\n\n' + ' '.join(job['id'])
        return formatted





class RosterDay:
    name_list = []
    epoch = None
    def __init__(self, roster, day):
        self.raw_roster = roster
        self.date = self.raw_roster.worksheets[0]['G1'].value + timedelta(days=day)
        self.day_sheet = roster.worksheets[day]
        #all dates must be calculated relative to the dat on sheet 0:
        self.shifts = []
        if not RosterDay.name_list:
            RosterDay.name_list = self.name_positions()
        if not RosterDay.epoch:
            RosterDay.epoch = input('what is epoch date?')

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
            job['id'] = self.get_cell_data(self.day_sheet, i, 'E') if self.get_cell_data(self.day_sheet, i, 'E') else None
            job['down'] = self.get_cell_data(self.day_sheet, i, 'F') if self.get_cell_data(self.day_sheet, i, 'F')  else None
            job['up'] = self.get_cell_data(self.day_sheet, i, 'H') if self.get_cell_data(self.day_sheet, i, 'H') else None
            job['down dest'] = self.get_cell_data(self.day_sheet, i, 'G') if self.get_cell_data(self.day_sheet, i, 'G') else None
            job['up dest'] = self.get_cell_data(self.day_sheet, i, 'I') if self.get_cell_data(self.day_sheet, i, 'I') else None
            job['start'] = int(self.get_cell_data(self.day_sheet, i, 'C')[0]) if self.get_cell_data(self.day_sheet, i, 'C') else None
            job['finish'] = int(self.get_cell_data(self.day_sheet, i, 'D')[0]) if self.get_cell_data(self.day_sheet, i, 'D') else None
            job['isrest'] = self.get_cell_data(self.day_sheet, i, 'J')[0] if self.get_cell_data(self.day_sheet, i, 'J') else False
            job['required'] = False if job['up'] and job['up'][0] in ['OFF', 'EDO'] else True
            job['debug'] = [self.day_sheet.title, i]
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
                break
        logging.debug('last cell is' + str(last_cell))
        return  last_cell.row + 1


    def name_positions(self):
        names = []
        last_cell = self.get_last_cell()
        for i in range(5,last_cell,4):
            names.append(str(self.get_cell_data(self.day_sheet, i, 'B')))
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
    print(location)
    #returns the most recent file with extention
    path = location + '\\*' + extention
    list_of_files = glob.glob(path)  # * means all if need specific format then *.csv
    print(list_of_files)
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file



class live_roster:
    def __init__(self, roster):
        self.master_roster = roster
        self.name = "Macalister"
        self.generated_roster = self.master_roster.generate_roster(self.name,48)
        self.swaps = []
        self.report_num = 0

    def swap_day(self, person, date, line):
        #accepts date object
        day_index = self.since_epoch(date) % 28
        days_since_epoch = self.since_epoch(date)
        self.generated_roster[days_since_epoch] = self.master_roster.compiled_roster[day_index].shifts[line-1]
        self.generated_roster[days_since_epoch]['date'] = date
        self.swaps.append([person, date, line])



    def swap_days(self, person, dates, line):
        #dates is a tuple first date and last date
        for day in dates:
            self.swap_day(person, day, line)

    def since_epoch(self, date):
        if type(date) == datetime:
            date = date.date()
        return (date - self.master_roster.epoch.date()).days

    def pay_report(self, period_ending, slip):
        period_beginning = period_ending - timedelta(days=13)
        worked_dates = [period_beginning + timedelta(days=x) for x in range(14)]
        target_shifts = [s for s in self.generated_roster if s['date'] in worked_dates]
        report = {}
        report['period ending'] = target_shifts[-1]['date']
        report['hrsworked'] = self.hours_worked(target_shifts)
        report['hrsearned'] = self.hours_earned(target_shifts)
        report['slipeworked'] = slip.advice['hours_worked']
        report['slipearned'] = slip.advice['hours_earned']
        return report

    def hours_worked(self, target_shifts):
        print("NUM OF SHIFTS:")
        print(target_shifts)
        [print(i['finish']) for i in target_shifts]
        wrked = timedelta()
        for i in target_shifts:
            print(wrked)
            if i['finish']:
                start, finish = self.dt_hours(i['start'], i['finish'])
                if finish.hour in [0,1]:
                    finish += timedelta(days=1)
                t = finish - start
                wrked += t
                continue
            if i['date'].date() in public_hols:
                wrked += timedelta(hours=8)
        return self.deltaToHrs(wrked)

    def dt_hours(self, start, finish):
        format = '%H%M'
        s = datetime.strptime(str(start).zfill(4), format)
        f = datetime.strptime(str(finish).zfill(4), format)
        return s, f

    def hours_earned(self, target_shifts):
        hrs_earned = timedelta()
        x = 0
        for i in target_shifts:
            if i['finish']:
                # calculate days pay
                start, finish = self.dt_hours(i['start'], i['finish'])
                dp = finish - start
                if i['date'].weekday() == 5:
                    dp *= 1.5
                if i['date'].weekday() == 6:
                    dp *= 2
                if i['date'].date() in public_hols:
                    dp *= 2.5
                hrs_earned += dp

            if i['up'] == ['EDO']:
                hrs_earned += timedelta(hours=8)
            x+=1
            #print(hrs_earned)

        return self.deltaToHrs(hrs_earned)



    @staticmethod
    def deltaToHrs(delta):
        hrs = delta.days * 24
        hrs += delta.seconds / 60 / 60
        return hrs













def unpickle():
    print(working_dir)
    if os.path.exists(working_dir+'\\'+'roster.pickle'):
        with open(working_dir+'\\'+'roster.pickle', 'rb') as file:
            roster =  pickle.load(file)
    else:
        roster =  Roster(working_dir)
        roster.roster_build()

    print(roster)

    return roster




def main():
    r = Roster(working_dir+'\\MasterRoster.pdf', pdfFullPages, pdfLastCoords, epoch)
    print(r.df.pivot(index='rosLine', columns='rosDay', values=['start','finish']).swaplevel(axis=1).sort_index(axis=1, ascending=[1,0]))
    print(r.df.unstack())



    '''
    with open(working_dir+'\\'+'roster.pickle', 'wb') as file:
            pickle.dump(roster, file)
    '''

if __name__ == "__main__":
    main()



# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

# TODO: fix XL.pickle and roster.pickle
# TODO: add py.test tests
# TODO: fix calander api max call errors
# TODO: add functionality to reconcile reported hours worked vs reported hours earned