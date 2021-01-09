#! python3
# slip.py - retrieves pdf from adp and then scrapes data from that pdf
#TODO: when updating roster you must input the number of full pages below as well as the coords for
# final page
# The roster must also be cleaned of all blue "available" this was done using 'Libre Draw'
pdfFullPages = 9
pdfLastCoords = [148.6, 107.0 ,1166.4, 251.5]
epoch = '17/01/21' #date in 30/6/99 format
import time, os, tabula,logging, shutil, holidays
import pandas as pd
from datetime import datetime, timedelta
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


    ## PANDAS DATA FRAME ??
class Roster:
    def __init__(self, roster_location, pdfFullPages, pdfLastCoords, epoch):
        self.roster_location = roster_location
        self.fix_cell_helper = 0
        self.pdfFullPages = pdfFullPages
        self.epoch = epoch
        self. pdfLastCoords = pdfLastCoords
        if os.path.exists(working_dir+"//rosterDataFrame.csv"):
            self.df = pd.read_csv(working_dir+"//rosterDataFrame.csv")
        else:
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
        df.to_csv(working_dir+"//rosterDataFrame.csv")

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




def main():
    r = Roster(working_dir+'\\MasterRoster.pdf', pdfFullPages, pdfLastCoords, epoch)
    view = r.df.pivot(index='rosLine', columns='rosDay', values=['start','finish']).swaplevel(axis=1).sort_index(axis=1, ascending=[1,0])
    print(view)



    '''
    with open(working_dir+'\\'+'roster.pickle', 'wb') as file:
            pickle.dump(roster, file)
    '''

if __name__ == "__main__":
    main()



# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# TODO: add py.test tests
# TODO: fix calander api max call errors
# TODO: add functionality to reconcile reported hours worked vs reported hours earned