import PySimpleGUIQt as sg
import slip, os, pickle
from datetime import datetime
from calendar import monthrange
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

def days_in_month(roster, month, year):
    days_in = []
    for d in roster.generated_roster:
        if d['date'].month == month:
                #and d['date'].year == datetime.now().year:
            days_in.append(d)
    return days_in

def fake_date():
    pass
fake_date.day = 'NA'



def change_month(roster, month, year):
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
    print(month)
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
    print(next_month)
    [print(i['date']) for i in next_month]
    for i in range(initial_weekday+len(roster_month),42):
        shift = next_month[calender_index]
        button_text = "{0}\n\n{1}".format(shift['date'].day, shift['start'])
        window[enc_btn(i)].update(button_text,button_color=('white', 'grey'))
        calender_index +=1




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

    layout = [
        [sg.Button('Previous', key="-PREV-"), sg.Text('JUN', justification='center', font=ff, key='-MONTH-'), sg.Button('Next', key='-NEXT-')],
        [Txt_day('Sun'), Txt_day('Mon'), Txt_day('Tue'), Txt_day('Wed'), Txt_day('Thu'), Txt_day('Fri'),
         Txt_day(' Sat')],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day(), Btn_day()],
        [sg.Button('Show'), sg.Button('Exit')]]

    window = sg.Window(date, layout)
    window.finalize()
    current_month = datetime.now().month
    current_year = datetime.now().year
    change_month(my_roster, current_month, datetime.now().year)


    while True:  # Event Loop
        event, values = window.read()
        print(event[:5], values)
        if event in (None, 'Exit'):
            break
        if event == 'Show':
            print('ANB')
            # Update the "output" text element to be the value of "input" element
            # window['-OUTPUT-'].update(values['-IN-'])
        if event == '-NEXT-':
            current_month,current_year = ((current_month +1,current_year) if current_month <12 else (1,current_year+1))
            change_month(my_roster, current_month, current_year)
        if event == '-PREV-':
            year = datetime.now().year
            if master_roster.epoch.month < datetime(year, current_month, 1).month:
                current_month, current_year = ((current_month - 1, current_year) if current_month > 2 else (12, current_year -1))
                change_month(my_roster, current_month, current_year)
        if event[:4] == "--XD":
            if not 'day1' in locals():
                print(event[4:-2])
                day1 = datetime(current_year, current_month, int(event[4:-2]))
            elif not 'day2' in locals():
                day2= datetime(current_year, current_month, int(event[4:-2]))
            print(day1)


    window.close()


    with open(working_dir+'\\'+'master_roster.pickle', 'wb') as file:
        pickle.dump(master_roster,file)
    with open(working_dir+'\\'+'my_roster.pickle', 'wb')  as file:
        pickle.dump(my_roster,file)


