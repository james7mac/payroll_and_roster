import PySimpleGUIQt as sg
import roster, os, calendar, json
from datetime import datetime
from calendar import monthrange
from datetime import timedelta
import pandas as pd
import pandas as pdpd
from googlecal import update_calander, check_work_event, delete_event, get_creds



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
        return datetime(*date_tup).date()

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
    def __init__(self):
        self.swaps = []

    def add(self, swap_with, swap_to_line, swap_dates):
        print(type(swap_dates[0]))
        self.swaps.append({'swap_with': swap_with, 'to_line': swap_to_line, 'dates': swap_dates})

    def formatted_swaps(self):
        formatted = []
        for k,i in enumerate(self.swaps):
            print(i['dates'][0])
            dates = ' '.join([x.strftime('%d/%m,') for x in i['dates']])
            entry = str(k+1).ljust(4) + i['swap_with'] + i['to_line'].rjust(18) + dates.rjust(40)
            formatted.append(entry)
        return formatted

    def remove(self, list_position):
        itm = self.swaps[list_position-1]

        del self.swaps[list_position-1]
        #change_month(my_roster, current_month, current_year, selector)

def change_month(cal, selector):
    window['-YEAR-'].update(str(cal.date.year))
    window['-MONTH-'].update(calendar.month_name[cal.date.month])
    shade_swaps = []
    '''
    for i in swaps.swaps:
        print(shade_swaps)
        [shade_swaps.append(x.date())for x in i['dates']]
    '''
    shifts = get_shifts(cal, roster, line)
    for i, k in enumerate(cal.itermonthdates(cal.date.year, cal.date.month)):
        colors=sg.theme_button_color()
        text = 'off'
        print(k)
        try:
            print(swaps.swaps[0]['dates'][0]==k['dates'])
        except:
            pass

        if k in [i['dates'] for i in swaps.swaps]:
            print('THERE IS A SWAP')
        if not pd.isnull(shifts.iloc[i].start):
            text = '{}:{}'.format(shifts.iloc[i].start.hour, shifts.iloc[i].start.minute)
        button_text = "{0}\n\n{1}".format(k.day,text)
        window[enc_btn(i)].update(button_text, button_color=(colors))
        #if d['date'].date() in shade_swaps:
        #    window[enc_btn(i + initial_weekday)].update(button_color=('yellow', 'navy'))



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

def get_shifts(cal, roster, line):
    daysSinceEpoch = (cal.date - roster.epoch).days
    line = line + daysSinceEpoch//28 % 83
    day = daysSinceEpoch%28
    if line  < 82:
        first = roster.df.index.get_loc((line, day+1))-1
        return roster.df[first:first+42]

    else:
        looping_df = roster.df[-56:].append(roster.df[:56])
        print(looping_df)
        first = looping_df.index.get_loc((line, day + 1)) - 1
        return roster.df[first:first + 42]








if __name__ == "__main__":

    roster = roster.Roster()
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
        [sg.Text('', size=[1,1])],
        [Txt('Swap with: '), In('name',key='-INSWAP-'), Txt('into line: '), In('0', key='-INLINE-')],
        [Btn_swap()],
        [sg.Listbox([],key='-SWAPLIST-')],
        [Txt('Delete Swap:'), In('', key='-DELSWAP-'), sg.Button('delete', key='-DELSWAPBTN-')],
        [sg.Button('Exit')]
    ]

    window = sg.Window(date, layout)
    window['-SWAPLIST-'].old_shifts = []
    window.finalize()


    if os.path.exists(working_dir+'//guiSettings.json'):
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


    if os.path.exists(working_dir+'//swaps.json'):
        with open(working_dir + "\\json.csv") as file:
            swap_dict = json.load(file)
        swaps = Swaps()
        swaps.swaps = swap_dict
    else:
        swaps = Swaps()


    buttonc0l = window['-SWAP-'].ButtonColor
    selector = select_date()
    #window['-SWAPLIST-'].update(swaps.formatted_swaps())
    cal = calend(datetime.now())
    #todo delete dummy line below
    cal.date = cal.date + timedelta(days=20)
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




        if event[:4] == "--XD":
            selector.date(cal, event)


        if event == "-SWAP-":
            swap_to_line = values['-INLINE-']
            swap_with = values['-INSWAP-']

            swap_dates = [selector.get_date(cal, x[0]) for x in selector.dates]
            swaps.add(swap_with, swap_to_line, swap_dates)
            selector.dates = []
            window['-SWAPLIST-'].update(swaps.formatted_swaps())
            change_month(cal, selector)
            cal_popup = sg.popup_yes_no("would you like to update google calendar?")
            if cal_popup == 'Yes':
                service = get_creds()
                new_shifts = [i for i in my_roster.generated_roster if i['date'] in swap_dates]
                master_roster.update_calander(new_shifts, service)

        if event == "-DELSWAPBTN-":
            swaps.remove(int(values['-DELSWAP-']))
            window['-SWAPLIST-'].update(swaps.formatted_swaps())


    window.close()



