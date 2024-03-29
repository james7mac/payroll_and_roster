import PySimpleGUIQt as sg
import roster, os, calendar, json, csv
from datetime import datetime
from datetime import date as DATE
from calendar import monthrange
from datetime import timedelta
import pandas as pd
from googlecal import update_calander, check_work_event, delete_event, get_creds
from sqlalchemy import create_engine


#show roster line

monkeypatch = False
startTimeWarning = 11
if os.environ['COMPUTERNAME'] == 'JMLAPTOP':
    working_dir =   r'C:\Users\james\PycharmProjects\payroll_and_roster'

elif os.environ['COMPUTERNAME'] == 'JAMESPC':
    working_dir = r'D:\Python\payroll_and_roster'
else:
    Exception()

def enc_btn(day):
    return '--XD' + str(day) + '--'

def Btn_day(*args, **kwargs):
    Btn_day.count += 1
    return (sg.Button(*args, "X", size_px=(70, 70), font=(f, 12),
                      key=(enc_btn(Btn_day.count-1)), **kwargs))
Btn_day.count = 0

def weekText():
    week = []
    for i in range(6):
        week.append(sg.Text("-", key=('-wk{}-'.format(str(i)))))
    return week
        
        

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
        if window[event].ButtonColor[1] == 'grey':
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


    for i,k in enumerate(cal.itermonthdays(cal.date.year, cal.date.month)):
        firstDay = i
        if k != 0:
            break

    rLine = str(shifts.iloc[firstDay].name[0])



    window['-rosterLine-'].update(rLine)


    for i, k in enumerate(cal.itermonthdates(cal.date.year, cal.date.month)):
        colors=sg.theme_button_color()
        window[enc_btn(i)].update(button_color=(colors))
    
        if not pd.isnull(shifts.iloc[i].start):
            text = shifts.iloc[i].start.strftime('%H:%M')

            if shifts.iloc[i].start.hour < startTimeWarning:
                window[enc_btn(i)].update(button_color=('white', 'maroon'))
        else:
            text = shifts.iloc[i].id




        #check for swaps in swap dataframe
        s= swaps.swaps

        if not s.empty:
            if not s[s.dates==pd.Timestamp(k)].empty:
                row = s[s.dates==pd.Timestamp(k)]
                SWAP = swap(row)
                if pd.notnull(SWAP.start):
                    text = SWAP.start.strftime('%H:%M')
                elif SWAP.id in ['OFF', 'EDO']:
                    text = SWAP.id
                else:
                    raise
                window[enc_btn(i)].update(button_color=('white', 'navy'))

        button_text = "{0}\n\n{1}".format(k.day, text)
        window[enc_btn(i)].update(button_text)

    for i,k in enumerate(cal.itermonthdays3(cal.date.year, cal.date.month)):
        month, day = k[1], k[2]
        if month != cal.date.month:
            button_text = "{0}\n\n{1}".format(day, '')
            window[enc_btn(i)].update(button_text, button_color=('white', 'grey'))


    for i in range(35,42):
        if len([j for j in cal.itermonthdays3(cal.date.year, cal.date.month)]) < 36:
            window[enc_btn(i)].update(visible=False)   # button_color=('white','white'))
        else:
            window[enc_btn(i)].update(visible=True)
    
    #update week label
    cal_month = list(cal.itermonthdays3(cal.date.year, cal.date.month))
    index = 0
    for i in range(0,42,7):
        if i >= len(cal_month):
            window['-wk{}-'.format(int(i/7))].update(' ')
            continue
        if (DATE(*cal_month[i]) - roster.epoch.date()).days//14%2==0:
            window['-wk{}-'.format(int(i/7))].update('1')
        else:
            window['-wk{}-'.format(int(i/7))].update('2')

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

def find_line(date, name):
    daysSinceEpoch = (date - roster.epoch).days
    initialLine = names.index(name.lower()) + 1
    print(daysSinceEpoch//28)
    return ((daysSinceEpoch// 28) + initialLine)%83

def find_name(name):
    if name in names:
        return name
    else:
        possible = []
        for i in names:
            if i.find(name) != -1:
                possible.append(i)
        new_try = sg.popup_get_text("Name Not Found. Use part of the name and it may show below:\n{}".format(
            str(possible)))
        if new_try:
            find_name(new_try.lower())
        else:
            return None


if __name__ == "__main__":
    os.chdir(working_dir)
    print(os.getcwd())
    if monkeypatch:
        roster.Roster.update_calander = patch_calendar
    roster = roster.Roster()
    totalRosterLines = len(roster.df.groupby(level=0))
    with open('roster_order') as file:
        names = file.read().split('\n')
        names = [i.lower() for i in names]
    date = "Today: " + datetime.now().strftime('%d/%m/%y')
    f = 'Helvetica'
    ff = (f, 14)
    sg.theme('LightGrey')

    #col1 = [sg.Listbox(values=['ITCH', 'NI', 'SUN'], size=(20,20), key='-SWAPLIST-')]
    wk0,wk1,wk2,wk3,wk4,wk5 = weekText()
    layout = [
        [sg.Button('Previous', key="-PREV-"), Txt_dt(key='-MONTH-', justification='r'), Txt_dt(key='-YEAR-', justification='l'), sg.Button('Next', key='-NEXT-')],
        [sg.Text(' '),Txt_day('Sun'), Txt_day('Mon'), Txt_day('Tue'), Txt_day('Wed'), Txt_day('Thu'), Txt_day('Fri'),
         Txt_day(' Sat')],
        [sg.Text('', key='-rosterLine-')],
        [wk0,Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [wk1,Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [wk2,Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [wk3,Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [wk4,Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [wk5,Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [sg.Button(' Upload Month', key='-uploadMonth-'),sg.Button('Find Line', key='-findLine-')],
        [sg.Text('', size=[1,1])],
        [Txt('Swap with: '), In('name',key='-INSWAP-'), Txt('into line: '), In('0', key='-INLINE-')],
        [Btn_swap()],
        [sg.Listbox([],key='-SWAPLIST-',font=('Courier',10))],
        [Txt('Delete Swap:'), In('', key='-DELSWAP-'), sg.Button('delete', key='-DELSWAPBTN-')],
        [sg.Button('Get Cowork Roster', key='-GETCOWORKER-')],
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
            s = swaps.swaps
            s = s[s.dates>(datetime.now()-timedelta(days=7))]
            print(s)
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
        
        if event == '-GETCOWORKER-':
            coworker = sg.popup_get_text("Who's roster are you looking for?")
            coworker = find_name(coworker)
            if coworker:
                #find line on original roster
                coLine = names.index(coworker.lower()) + 1
                print(coLine)
                coRost = get_shifts(cal, roster, coLine)
                coRost = add_dates(cal, coRost)
                coRost['date'] = coRost['date'].dt.strftime('%Y-%m-%d')
                coRost['start'] = coRost['start'].apply(lambda x: x.strftime('%H:%M:%S')if not pd.isnull(x) else '')
                coRost['finish'] = coRost['finish'].apply(lambda x: x.strftime('%H:%M:%S')if not pd.isnull(x) else '')
                coRost = coRost.filter(['date', 'start', 'finish', 'job', 'id'])#.set_index('date')
                print(coRost)
                
                df_layout = [
                [sg.Table(values=coRost.values.tolist(),
                                  headings=list(coRost.columns),
                                  pad=(5,5),
                                  display_row_numbers=False,
                                  auto_size_columns=True,
                                  num_rows=min(31, len(coRost.values.tolist()))
                                  )]
                ]
                window2 = sg.Window(coworker,df_layout)
                event, values = window2.read()
                event = 'None'
                window2.close()

                
            
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

        if event == '-findLine-':
            target = sg.popup_get_text('Who are you looking for?')
            if target:
                target.lower()
                valid_name= find_name(target)
                lines = []
                #print(selector.get_date(cal,selector.dates[0][0]))
                for i in selector.dates:
                    lines.append(find_line(selector.get_date(cal,i[0]), valid_name))
                sg.popup_ok('{} found on lines: {}'.format(valid_name, lines))

        if event[:4] == "--XD":
            selector.date(cal, event)

        
        if event == "-SWAP-":

            swap_to_line = values['-INLINE-']
            try:
                int(swap_to_line)
            except:
                sg.popup('swap line must be a number')
                swap_to_line = 0
            if int(swap_to_line) >= 1 and int(swap_to_line) <= totalRosterLines:
                swap_with = values['-INSWAP-']
                overwrite = True
                swap_dates = [selector.get_date(cal, x[0]) for x in selector.dates]
                swap_dates.sort()
                check_previous =  set(swap_dates).intersection(swaps.swaps.dates)
                if check_previous:
                    s = swaps.swaps
                    old_swaps = s[s.dates.isin(check_previous)]
                    print(old_swaps)
                    overwrite = sg.popup_yes_no('Overwrite old swaps with {}'.format([i for i in old_swaps.swap_with]))
                    if overwrite == 'Yes':
                        swaps.swaps = swaps.swaps[~swaps.swaps.dates.isin(check_previous)]
                    else:
                        overwrite = False
    
                for dayte in swap_dates:
                    if not overwrite:
                        break
                    if (dayte-roster.epoch).days//28 == (swap_dates[0]-roster.epoch).days//28:
                        swaps.add(swap_with, swap_to_line, dayte)
                    else:
                        swaps.add(swap_with, str(int(swap_to_line)+1), dayte)
                        print(swaps.swaps)
                swaps.swapID+=1
    
                swapcal = calend(datetime(swap_dates[0].year, swap_dates[0].month, 1))
                selector.dates = []
    
                window['-SWAPLIST-'].update(swaps.formatted_swaps())
                change_month(cal, selector)
            else:
                sg.popup('Roster Lines only valid from 1 to {}'.format(totalRosterLines))


            '''
            swapcal = calend(datetime(swap_dates[0].year, swap_dates[0].month, 1))
            change_month(cal, selector)
            cal_popup = sg.popup_yes_no("would you like to update google calendar?")
            if cal_popup == 'Yes' and overwrite == 'Yes':
                swapDF = get_shifts(swapcal, roster, int(swap_to_line))
                swapDF = add_dates(swapcal, swapDF)
                swapDF = swapDF[swapDF['date'].isin(swap_dates)]
                print(swapDF)
                roster.update_calander(swapDF)

            '''

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
        #second DB connection to creat a backup once per week
        con2 = create_engine("sqlite:///swaps_BACKUP.db", echo=False)
        if os.path.exists('swaps_BACKUP.db'):
           if os.path.getmtime('swaps_BACKUP.db') + timedelta(days=7).total_seconds() < datetime.today().timestamp():
                swaps.swaps.to_sql('Swaps', con2, if_exists='replace', index=False)
        else:
            swaps.swaps.to_sql('Swaps', con2, if_exists='replace', index=False)



    window.close()



