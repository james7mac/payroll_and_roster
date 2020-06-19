import PySimpleGUIQt as sg
import slip, os, pickle, json
from datetime import datetime
from calendar import monthrange
from collections import deque
from datetime import timedelta

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


month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

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

    def get_date(self, event):
        return datetime(current_year, current_month, int(self.get_tile_value(event)))

    def date(self, event):
        print(window[event].ButtonColor[1])
        if window[event].ButtonColor[1] != buttonc0l[1] and window[event].ButtonColor[1] != 'black':
            return
        month = month_names.index(window['-MONTH-'].DisplayText.upper()) + 1
        if (event, month) in self.dates:
            self.dates.remove((event, month))
            self.shade(month)
            return
        date = self.get_date(event)
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


class Swaps:
    def __init__(self):
        self.swaps = []

    def add(self, swap_with, swap_to_line, swap_dates, old_shifts):
        self.swaps.append({'swap_with': swap_with, 'to_line': swap_to_line, 'dates': swap_dates, 'old_shifts': old_shifts})
        my_roster.swap_days(swap_with, swap_dates, int(swap_to_line))

    def formatted_swaps(self):
        formatted = []
        for k,i in enumerate(self.swaps):
            dates = ' '.join([x.strftime('%d/%m/%y') for x in i['dates']])
            entry = str(k+1).ljust(4) + i['swap_with'] + i['to_line'].rjust(18) + dates.rjust(40)
            formatted.append(entry)
        return formatted

    def remove(self, list_position):
        itm = self.swaps[list_position-1]
        for x, day in enumerate(itm['dates']):
            for k, i in enumerate(my_roster.generated_roster):
                if i['date'] == day:
                    my_roster.generated_roster[k] = itm['old_shifts'][x]

        del self.swaps[list_position-1]




def change_month(roster, month, year, selector):

    roster_month = days_in_month(roster, month, year)
    prev_month = days_in_month(roster,month -1, year if month > 1 else days_in_month(roster,12,year-1))
    next_month = days_in_month(roster,month + 1,year) if month < 12 else days_in_month(roster,1, year+1)
    initial_weekday = roster_month[0]['date'].weekday() + 1
    window['-MONTH-'].update(roster_month[0]['date'].strftime('%b'))
    for i,d in enumerate(roster_month):
        if d['date'].date() == roster.master_roster.epoch.date():
            colors = ('white', 'red')
        else:
            colors=sg.theme_button_color()
        button_text = "{0}\n\n{1}".format(d['date'].day,d['start'])
        window[enc_btn(i+initial_weekday)].update(button_text, button_color=(colors))


    prev_month = [i for i in prev_month if i['date'].year == roster_month[0]['date'].year]

    prev_month = prev_month[-initial_weekday:]
    previous_month_num = month-1 if month != 1 else 12
    num_days_prev = monthrange(roster_month[0]['date'].year, previous_month_num)[1]
    remaining_days = num_days_prev - initial_weekday

    for i in range(initial_weekday):
        shift = (prev_month[i] if len(prev_month) > i else  {'date':fake_date,'start':'NA'})
        button_text = "{0}\n\n{1}".format(shift['date'].day, shift['start'])
        window[enc_btn(i)].update(button_text,button_color=('white','grey'))

    #length of calander grid minus start position minus length of month gives remamings days on grid
    remaining_spots = 42-initial_weekday-len(roster_month)
    next_month = next_month[:remaining_spots]
    calender_index = 0
    for i in range(initial_weekday+len(roster_month),42):
        shift = next_month[calender_index]
        button_text = "{0}\n\n{1}".format(shift['date'].day, shift['start'])
        window[enc_btn(i)].update(button_text,button_color=('white', 'grey'))
        calender_index +=1


    selector.shade(month)





if __name__ == "__main__":
    if os.path.exists(working_dir+'\\'+'master_roster.pickle'):
        with open(working_dir+'\\'+'master_roster.pickle', "rb") as file:
            master_roster = pickle.load(file)
    else:
        master_roster = slip.Roster(working_dir)
        master_roster.roster_build()
    if  os.path.exists(working_dir+'\\'+'my_roster.pickle'):
        with open(working_dir+'\\'+'my_roster.pickle','rb') as file:
            my_roster = pickle.load(file)
    else:
        my_roster = slip.live_roster(master_roster)


    date = "Today: " + datetime.now().strftime('%d/%m/%y')
    f = 'Helvetica'
    ff = (f, 14)
    sg.theme('LightGrey')

    col1 = [sg.Listbox(values=month_names, size=(20,20), key='-SWAPHIST-')]

    layout = [
        [sg.Button('Previous', key="-PREV-"), sg.Text('', justification='center', font=ff, key='-MONTH-'), sg.Button('Next', key='-NEXT-')],
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
    buttonc0l = window['-SWAP-'].ButtonColor
    current_month = datetime.now().month
    current_year = datetime.now().year
    selector = select_date()

    change_month(my_roster, current_month, datetime.now().year, selector)

    if os.path.isfile('swaps.pickle'):
        with open(working_dir + '\\' + 'swaps.pickle', 'rb') as file:
            swaps = pickle.load(file)
            window['-SWAPLIST-'].update(swaps.formatted_swaps())
    else:
        swaps = Swaps()

    while True:  # Event Loop
        event, values = window.read()
        print(event, values)
        if event in (None, 'Exit'):
            break

        if event == '-NEXT-':
            current_month,current_year = ((current_month +1,current_year) if current_month <12 else (1,current_year+1))
            change_month(my_roster, current_month, current_year, selector)

        if event == '-PREV-':
            year = datetime.now().year
            if master_roster.epoch.month < datetime(year, current_month, 1).month:
                current_month, current_year = ((current_month - 1, current_year) if current_month > 2 else (12, current_year -1))
                change_month(my_roster, current_month, current_year, selector)

        if event[:4] == "--XD":
            selector.date(event)

        if event == "-SWAP-":
            swap_to_line = values['-INLINE-']
            swap_with = values['-INSWAP-']
            swap_dates = [selector.get_date(x[0]) for x in selector.dates]
            old_shifts = [i for i in my_roster.generated_roster if i['date'] in swap_dates]
            swaps.add(swap_with, swap_to_line, swap_dates, old_shifts)
            window['-SWAPLIST-'].update(swaps.formatted_swaps())

        if event == "-DELSWAPBTN-":
            swaps.remove(int(values['-DELSWAP-']))
            window['-SWAPLIST-'].update(swaps.formatted_swaps())

            # TODO: ADD REMOVE SWAP FUNCTION

            # TODO: ADD UPDATE GOOGLE CAL FUNCTION

            # TODO: ADD SOME EARLY HIGHLIGHTING OR WARNING FEATURE




        window.refresh()

    window.close()


    with open(working_dir+'\\'+'master_roster.pickle', 'wb') as file:
        pickle.dump(master_roster,file)
    with open(working_dir+'\\'+'my_roster.pickle', 'wb')  as file:
        pickle.dump(my_roster,file)
    with open(working_dir + '\\' + 'swaps.pickle', 'wb') as file:
        pickle.dump(swaps, file)



