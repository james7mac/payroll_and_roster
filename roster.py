#! python3
# slip.py - retrieves pdf from adp and then scrapes data from that pdf
#TODO: when updating roster you must input the number of full pages below as well as the coords for
# final page
# The roster must also be cleaned of all blue "available" this was done using 'Libre Draw'
pdfFullPages = 9
pdfLastCoords = [148.6, 107.0 ,1166.4, 251.5]
nameCoords = []
epoch = '17/01/21' #date in 30/6/99 format
import time, os, tabula,logging, holidays
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
    working_dir = r'D:\Python\payroll_and_roster'

else:
    Exception()

class Roster:
    def __init__(self, roster_location=working_dir, pdfFullPages=pdfFullPages, pdfLastCoords=pdfLastCoords, epoch=epoch):
        self.roster_location = roster_location
        self.fix_cell_helper = 0
        self.pdfFullPages = pdfFullPages
        self.epoch = epoch
        self. pdfLastCoords = pdfLastCoords
        self.service = None
        if os.path.exists(working_dir+"//rosterDataFrame.csv"):
            self.df = pd.read_csv(working_dir+"//rosterDataFrame.csv", parse_dates=['start','finish'])
            self.df['start'] = pd.to_datetime(self.df['start']).dt.time
            self.df['finish'] = pd.to_datetime((self.df['finish'])).dt.time
        else:
            self.df = self.roster_build(self.roster_location+'\\MasterRoster.pdf', pdfFullPages, pdfLastCoords)
        self.df = self.df.set_index(['rosLine','rosDay'])




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



