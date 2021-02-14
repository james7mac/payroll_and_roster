import PySimpleGUIQt as sg
import roster, os, calendar, json, csv
from datetime import datetime
from calendar import monthrange
from datetime import timedelta
import pandas as pd
from googlecal import update_calander, check_work_event, delete_event, get_creds
from sqlalchemy import create_engine

monkeypatch = False

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

def enc_btn(day):
    return '--XD' + str(day) + '--'

def Btn_day(*args, **kwargs):
    Btn_day.count += 1
    return (sg.Button(*args, "X", size_px=(70, 70), font=(f, 12),
                      key=(enc_btn(Btn_day.count-1)), **kwargs))
Btn_day.count = 0

def Txt_day(*args, **kwargs):
    return (sg.Text(*args, size_px=(70, 30), font=ff, justification='center', background_color='lightgrey',
                    **kwargs))
def Btn_swap(*args, **kwargs):
    return (sg.Button(*args, "Swap",size_px=(520, 30), font=(f, 12), key=('-SWAP-'), **kwargs))
def In(*args, **kwargs):
    return (sg.InputText(*args, font=(f, 12), **kwargs))
def Txt(*args, **kwargs):
    return (sg.Text(*args, font=(f, 12), **kwargs))
def Txt_dt(*args, **kwargs):
    return (sg.Text('', font=ff, **kwargs))

def days_in_month(roster, month, year):
    days_in = []
    for d in roster.generated_roster:
        if d['date'].month == month:
                #and d['date'].year == datetime.now().year:
            days_in.append(d)
    return days_in

def date_in_roster(roster, date):
    for d in roster:
        if d['date'] == date:
            return date

def fake_date():
    pass
fake_date.day = 'NA'


class select_date:
    def __init__(self):
        self.dates = []
        self.year = {}

    def get_tile_value(self, event):
        return event[4:-2]

    def get_date(self, cal, event):
        tile = int(self.get_tile_value(event))
        date_tup = [i for i in cal.itermonthdays3(cal.date.year, cal.date.month)][tile]
        return datetime(*date_tup)

    def date(self,cal, event):
        if window[event].ButtonColor[1] != buttonc0l[1] and window[event].ButtonColor[1] != 'black':
            return
        month = cal.date.month
        if (event, month) in self.dates:
            self.dates.remove((event, month))
            self.shade(month)
            return
        self.dates.append((event, month))
        self.shade(month)


    def shade(self, month_name):
        month = range(0,41)
        if month_name not in self.year:
            self.year[month_name] = []
            for t in month:
                self.year[month_name].append((window[enc_btn(t)].ButtonColor))

        for tile in month:
            window[enc_btn(tile)].update(button_color=('white', self.year[month_name][tile][1]))

        for selected in self.dates:
            if month_name == selected[1]:
                window[selected[0]].update(button_color=('white', 'black'))




class calend(calendar.Calendar):
    def __init__(self, date, firstweekday=6):
        super().__init__(firstweekday=firstweekday)
        self.date = date

    def next_month(self):
        if self.date.month < 12:
            self.date = datetime(self.date.year, self.date.month +1, 1)
        else:
            self.date = datetime(self.date.year+1,1, 1)
        return self.date

    def prev_month(self):
        if self.date.month <= roster.epoch.month:
            return
        if self.date.month > 1:
            self.date = datetime(self.date.year, self.date.month -1, 1)
        else:
            self.date = datetime(self.date.year-1,1, 1)
        return self.date



class Swaps:
    swapID = 0
    def __init__(self):
        self.swaps = pd.DataFrame()

    def add(self, swap_with, swap_to_line, swap_dates):
        self.swaps = self.swaps.append(pd.DataFrame({'swapID':[self.swapID],'swap_with': [swap_with], 'to_line': [swap_to_line],\
                                                     'dates': [swap_dates]}), ignore_index=True)


    def formatted_swaps(self):
        formatted = []
        swaplist = self.swaps.sort_values(by='dates')
        for i,row in swaplist.iterrows():
            entry = "{}{}{}id: {}".format(row['dates'].strftime('%d/%m').ljust(8), row['swap_with'].ljust(20), row['to_line'].ljust(6), str(row['swapID']).ljust(3))

            formatted.append(entry)
        return formatted

    def remove(self, list_position):
        itm = self.swaps[list_position-1]

        del self.swaps[list_position-1]
        #change_month(my_roster, current_month, current_year, selector)

def delete_google_cal_between(date1, date2):
    for i in pd.date_range(date1, date2):
        service = get_creds()
        existing_event = check_work_event(i,service)
        if existing_event:
            delete_event(existing_event['id'], service)


def swap(swp):
    date, to_line = swp['dates'], int(swp['to_line'].values)
    df = roster.df
    df = df.loc[to_line]
    epochDays = int((date-roster.epoch).dt.days)
    column = (epochDays % 28) + 1
    shift = df.loc[column]
    shift['rosLine'] = to_line
    shift['rosDay'] = column
    shift['date'] = date
    return shift


def change_month(cal, selector):
    window['-YEAR-'].update(str(cal.date.year))
    window['-MONTH-'].update(calendar.month_name[cal.date.month])
    shade_swaps = []

    shifts = get_shifts(cal, roster, line)




    for i, k in enumerate(cal.itermonthdates(cal.date.year, cal.date.month)):
        colors=sg.theme_button_color()
        text = 'off'
        if not pd.isnull(shifts.iloc[i].start):
            text = shifts.iloc[i].start.strftime('%H:%M')


        #check for swaps in swap dataframe
        s= swaps.swaps
        if not s.empty:
            if not s[s.dates==pd.Timestamp(k)].empty:
                print(pd.Timestamp(k))
                row = s[s.dates==pd.Timestamp(k)]
                SWAP = swap(row)
                if pd.notnull(SWAP.start):
                    text = SWAP.start.strftime('%H:%M')
                elif SWAP.id in ['OFF', 'EDO']:
                    text = 'off'
                else:
                    raise

        button_text = "{0}\n\n{1}".format(k.day, text)
        window[enc_btn(i)].update(button_text, button_color=(colors))

    for i,k in enumerate(cal.itermonthdays3(cal.date.year, cal.date.month)):
        month, day = k[1], k[2]
        if month != cal.date.month:
            button_text = "{0}\n\n{1}".format(day, '')
            window[enc_btn(i)].update(button_text, button_color=('white', 'grey'))


    for i in range(35,42):
        if len([i for i in cal.itermonthdays3(cal.date.year, cal.date.month)]) < 36:
            window[enc_btn(i)].update(visible=False)   # button_color=('white','white'))
        else:
            window[enc_btn(i)].update(visible=True)


def apply_months_swaps(month):
    #must have add_dates() applied first
    SWAPS = swaps.swaps[swaps.swaps['dates'].isin(month.date)]
    for i, k in month.iterrows():
        for o, p in SWAPS.iterrows():
            if k.date == p.dates:
                new_shift = swap(pd.DataFrame(p).T)
                month.at[i, 'start'] = new_shift.start
                month.at[i, 'finish'] = new_shift.finish
                month.at[i, 'id'] = new_shift.id
                month.at[i, 'hours'] = new_shift.hours
                month.at[i, 'rest'] = new_shift.rest
    return month

def get_shifts(cal, roster, line):
    print(cal.date)
    daysSinceEpoch = (cal.date - roster.epoch).days
    running_line = (line-1 + (daysSinceEpoch//28)) % totalRosterLines + 1
    print('THIS IS LINE {}'.format(line))
    day = daysSinceEpoch%28
    print('total roster line.... {}'.format(running_line))
    if running_line  < totalRosterLines-1:
        first = roster.df.index.get_loc((running_line, day+1))-1
        print(first)
        print(roster.df[first:first+42])
        print('{} since epoch'.format(daysSinceEpoch))
        print('{} is day num'.format(day))
        print(calendar.monthrange(cal.date.year, cal.date.month))
        #subract the number of days at the beginning of calendar that are lead-in from prev month
        first-=calendar.monthrange(cal.date.year, cal.date.month)[0]


    else:
        looping_df = roster.df[-56:].append(roster.df[:56])
        first = looping_df.index.get_loc((running_line, day + 1)) - 1
    #hacky solution to glitch that occurs when first day of month is Sunday, causing roster
    #to be off by a week. May cause issue if a Sunday falls on the first at the same time as the
    #roster loops back around
    if cal.date.weekday() == 6:
        first+=7

    return roster.df[first:first + 42]


def add_dates(cal, shifts):
    # calDays has list of days in this month with a 0 for days from last month or next month
    calDays = [i for i in cal.itermonthdays(cal.date.year, cal.date.month)]
    monthdatelist = []
    for i, row in enumerate(shifts.itertuples()):
        #all days in calDays that arent this month are 0 and should be skipped
        if i >= len(calDays):
            break
        if calDays[i] == 0:
            continue

        monthdatelist.append(datetime(cal.date.year, cal.date.month, calDays[i]))

    #get rid of shifts not in this calender month
    x = [True if i > 0 else False for i in cal.itermonthdays(cal.date.year, cal.date.month)]
    shifts = shifts[:len(x)]
    shifts = shifts[x]
    #add dates to shifts datafame
    shifts['date'] = monthdatelist
    return shifts


def patch_calendar(self, jobs):
    print(jobs)
    for i, job in jobs.iterrows():
        if job['id'] not in ['OFF', 'EDO']:
            print('{}   on:  {}'.format(job['date'].strftime("%a %d-%m"), job['start']))
            #print(self.format_job(job))
        else:
            print('{}......NO WORK'.format(job['date']))



if __name__ == "__main__":
    if monkeypatch:
        roster.Roster.update_calander = patch_calendar
    roster = roster.Roster()
    totalRosterLines = len(roster.df.groupby(level=0))
    date = "Today: " + datetime.now().strftime('%d/%m/%y')
    f = 'Helvetica'
    ff = (f, 14)
    sg.theme('LightGrey')

    col1 = [sg.Listbox(values=['ITCH', 'NI', 'SUN'], size=(20,20), key='-SWAPHIST-')]

    layout = [
        [sg.Button('Previous', key="-PREV-"), Txt_dt(key='-MONTH-', justification='r'), Txt_dt(key='-YEAR-', justification='l'), sg.Button('Next', key='-NEXT-')],
        [Txt_day('Sun'), Txt_day('Mon'), Txt_day('Tue'), Txt_day('Wed'), Txt_day('Thu'), Txt_day('Fri'),
         Txt_day(' Sat')],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [sg.Button('Upload Month', key='-uploadMonth-')],
        [sg.Text('', size=[1,1])],
        [Txt('Swap with: '), In('name',key='-INSWAP-'), Txt('into line: '), In('0', key='-INLINE-')],
        [Btn_swap()],
        [sg.Listbox([],key='-SWAPLIST-',font=('Courier',10))],
        [Txt('Delete Swap:'), In('', key='-DELSWAP-'), sg.Button('delete', key='-DELSWAPBTN-')],
        [sg.Button('Exit')]
    ]

    window = sg.Window(date, layout)
    window['-SWAPLIST-'].old_shifts = []
    window.finalize()

    if datetime.now() > datetime(2021,2,1):
        cal = calend(datetime(datetime.now().year, datetime.now().month, 1))
    else:
        cal = calend(datetime(2021,2,1))

    if os.path.exists(working_dir+'\\guiSettings.json'):
        with open(working_dir + "\\guiSettings.json") as file:
            settings = json.load(file)
        line = settings['initialLine']

    else:
        settings={}
        settings['name'] = sg.popup_get_text('Please type your name then press ok')
        settings['initialLine'] = int(sg.popup_get_text('Ignoring all your swaps, what line are you on?'))
        with open(working_dir+"\\guiSettings.json",'w') as file:
            json.dump(settings,file)
        line = settings['initialLine']
        cal_popup = sg.popup_yes_no("would you like to update google calendar?")
        if cal_popup == 'Yes':
            service = get_creds()
            shifts=get_shifts(roster, line)
            #master_roster.create_calendar_event()


    con = create_engine("sqlite:///swaps.db", echo=False)

    if os.path.exists(working_dir+'\\swaps.db'):
        with open(working_dir + "\\swaps.db") as file:
            swaps = Swaps()
            print("loading swaps database...")
            swaps.swaps = pd.read_sql('Swaps', con)
            print(swaps.swaps)
            window['-SWAPLIST-'].update(swaps.formatted_swaps())
        #get last swapID from last sql entrry
        if not swaps.swaps.empty:
            swaps.swapID = swaps.swaps.iloc[-1].swapID + 1
        else:
            swaps = Swaps()
    else:
        swaps = Swaps()


    buttonc0l = window['-SWAP-'].ButtonColor
    selector = select_date()
    #window['-SWAPLIST-'].update(swaps.formatted_swaps())

    change_month(cal, selector)


    while True:  # Event Loop
        event, values = window.read()
        print(event, values)
        if event in (None, 'Exit'):
            break
        window.refresh()


        if event == '-NEXT-':
            cal.next_month()
            change_month(cal, selector)
            get_shifts(cal,roster,line)


        if event == '-PREV-':
            cal.prev_month()
            change_month(cal, selector)

        if event == '-uploadMonth-':
            shifts = get_shifts(cal, roster, line)
            month = add_dates(cal,shifts)
            month = apply_months_swaps(month)
            roster.update_calander(month)



        if event[:4] == "--XD":
            selector.date(cal, event)


        if event == "-SWAP-":
            swap_to_line = values['-INLINE-']
            swap_with = values['-INSWAP-']

            swap_dates = [selector.get_date(cal, x[0]) for x in selector.dates]
            print(swap_dates)
            for dayte in swap_dates:
                swaps.add(swap_with, swap_to_line, dayte)
            swaps.swapID+=1
            swapcal = calend(datetime(swap_dates[0].year, swap_dates[0].month, 1))
            selector.dates = []

            window['-SWAPLIST-'].update(swaps.formatted_swaps())




            change_month(cal, selector)
            cal_popup = sg.popup_yes_no("would you like to update google calendar?")
            if cal_popup == 'Yes':
                swapDF = get_shifts(swapcal, roster, int(swap_to_line))
                swapDF = add_dates(swapcal, swapDF)
                swapDF = swapDF[swapDF['date'].isin(swap_dates)]
                print(swapDF)
                roster.update_calander(swapDF)



        if event == "-DELSWAPBTN-":
            swapid = int(values['-DELSWAP-'])
            swaps.swaps = swaps.swaps[swaps.swaps.swapID != swapid]
            window['-SWAPLIST-'].update(swaps.formatted_swaps())
            cal_popup = sg.popup_yes_no("would you like to update google calendar?")

            # TODO: this code is cooked
            if cal_popup == 'Yes':
                service = get_creds()
                new_shifts = None
                master_roster.update_calander(new_shifts, service)

    if not swaps.swaps.empty:
        swaps.swaps.to_sql('Swaps', con, if_exists='replace', index=False)

    window.close()



