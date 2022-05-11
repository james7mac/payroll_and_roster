#! python3
# slip.py - retrieves pdf from adp and then scrapes data from that pdf
#TODO: when updating roster you must input the number of full pages below as well as the coords for
# final page
# The roster must also be cleaned of all blue "available" this was done using 'Libre Draw'

# TODO: the following coords are of unit 'Toporgraphic-points' and can be found easily in GIMP
pdfLastCoords = [101,148.6 ,299,1166.4]
nameCoords = []
epoch = '01/01/22' #date in 30/06/99 format
import time, os, tabula,logging, holidays, openpyxl
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

elif os.environ['COMPUTERNAME'] == 'JAMESPC':
    sheet_path = r'D:\Python\payroll_and_roster\seymourRoster.xlsm'

else:
    Exception()

roster_length = 11


class Roster:
    def __init__(self, sheet=sheet_path, epoch_=epoch):
        self.sheet = sheet
        self.epoch = epoch_
        self.df = self.create_dfs()

    def read_cell(self, raw_roster, coord):
        data = raw_roster.iloc[coord[0]:coord[0] + 5, coord[1]:coord[1] + 2]
        job = {}
        job['trains'] = ''
        if type(data.iloc[2, 0]) == str:
            job['trains'] = data.iloc[2, 0]
        if type(data.iloc[1, 1]) == str:
            job['trains'] = job['trains'] + ' ' + data.iloc[1, 1]

        if job['trains'] in ['Rostered Off', 'EDO', 'BOWP']:
            job['start'] = None
            job['finish'] = None
            job['job'] = None
        else:
            job['start'], job['finish'] = data.iloc[0, 0], data.iloc[0, 1]
            job['job'] = data.iloc[1, 0]
        return job

    def cell_coord(self, line, day):
        return [(line - 1) * 5, (day - 1) * 2]

    def create_df(self, week):
        name = "PROPOSED ROSTER WEEK " + str(week)
        raw_roster = pd.read_excel(self.sheet, sheet_name=name, header=None)
        roster_df = pd.DataFrame()
        for line in range(1, roster_length + 1):
            for day in range(1, 8):
                job = self.read_cell(raw_roster, self.cell_coord(line, day))
                job['day'] = day
                job['line'] = line
                job['week'] = week
                df = pd.DataFrame([job])
                roster_df = pd.concat([roster_df, df])
        return roster_df

    def create_dfs(self):
        dfs = []
        for i in range(1, 5):
            dfs.append(self.create_df(i))
        df = pd.concat(dfs)
        return df







'''
    def roster_build(self, pdfPath, pdfFullPages, pdfLastCoords):
        df1 = pd.DataFrame()
        for i in range(1, pdfFullPages+1):
            roster_pdf = open(pdfPath, 'rb')
            # coords of the area in the roster found using GIMP:

            #if i == pdfFullPages+1:
            #    x = tabula.read_pdf(roster_pdf, pages=int(i), lattice=1, area=pdfLastCoords)
            #else:
            x = tabula.read_pdf(roster_pdf, pages=int(i), lattice=1, area=[101, 149, 591, 1167.9])
            if type(x) == list:
                x = x[0]
            x.dropna(axis=0, thresh=30, inplace=True)
            x.dropna(axis=1, how='all', inplace=True)
            x.columns = [str(i) for i in range(1, 15)] + ['h1'] + [str(i) for i in range(15, 29)] + ['h2']
            df1 = df1.append(x)
        df1.index = [str(i + 1) for i in range(len(df1))]
        df1.drop(['h1', 'h2'], axis=1, inplace=True)
        print(df1)
        df1 = df1.applymap(self.fix_cell)
        df = pd.DataFrame()
        for index, i in enumerate(df1.iterrows()):
            for ii, dic in enumerate(i[1]):
                print('dic is')
                print(dic)
                print('row is {}, col is {}'.format(index, ii))
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
        update_calander(self.format_job(job), service)
        logging.debug('CAL UPDATE: ' + str(job['date']))

    def update_calander(self, jobs):
        if not self.service:
            self.service = get_creds()
        for i,job in jobs.iterrows():
            existing_event = check_work_event(job['date'], self.service)
            if existing_event:
                dstring = existing_event['start']['dateTime'][:-6]
                d = datetime.strptime(dstring,'%Y-%m-%dT%H:%M:%S')
                print('deleting event {}'.format(d))
                delete_event(existing_event['id'], self.service)

            if job['id'] in ['OFF', 'EDO']:
                existing_event = check_work_event(job['date'], self.service)
                if existing_event:
                    delete_event(existing_event['id'], self.service)



            if job['id'] not in ['OFF', 'EDO']:
                print('EVENT BEING CREATED')
                self.create_calander_event(job, self.service)


    def format_job(self, job):
        for i,k in job.items():
            print("{} is {}".format(i,k))
        if job['id'] in ['OFF','EDO']:
            return None
        formatted = {}
        if not job['rest']:
            formatted['summary'] = "Work"
        else:
            formatted['summary'] = 'Work - Rest'

        start = datetime.combine(job['date'],job['start'])
        end = datetime.combine(job['date'], job['finish'])
        if end.hour in [0,1]:
            end = end + timedelta(days=1)
        formatted['start'] = {'dateTime': start.isoformat(), 'timeZone':'Australia/Melbourne'}
        formatted['end'] = {'dateTime': end.isoformat(), 'timeZone':'Australia/Melbourne'}
        formatted['description'] =job['id']
        return formatted



'''
def main():
    r = Roster(sheet_path,epoch)
    r
    #view = r.df.pivot(index='rosLine', columns='rosDay', values=['start','finish']).swaplevel(axis=1).sort_index(axis=1, ascending=[1,0])
    print(r)



    '''
    with open(working_dir+'\\'+'roster.pickle', 'wb') as file:
            pickle.dump(roster, file)
    '''

if __name__ == "__main__":
    main()



