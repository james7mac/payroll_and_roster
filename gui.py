import PySimpleGUI as sg
import slip
from datetime import datetime

date = "Today: " + datetime.now().strftime('%d/%m/%y')

def Btn_day(*args, **kwargs):
    return (sg.Button(*args, size=(6,1),font=("Helvetica", 15), **kwargs))


def Txt_day(*args, **kwargs):
    return (sg.Text(*args, size=(6,1), font=("Helvetica", 15), **kwargs))

#sg.theme('BluePurple')


layout = [
          [Txt_day('Sun'), Txt_day('Mon'), Txt_day('Tue'), Txt_day('Wed'), Txt_day('Thu'), Txt_day('Fri'), Txt_day(' Sat')],
          [Btn_day('1'), Btn_day('1'), Btn_day('1'), Btn_day('1'), Btn_day('1'), Btn_day('1'), Btn_day('1')],
          [sg.Button('Show'), sg.Button('Exit')]]

window = sg.Window(date, layout)




while True:  # Event Loop
    event, values = window.read()
    print(event, values)
    if event in  (None, 'Exit'):
        break
    if event == 'Show':
        # Update the "output" text element to be the value of "input" element
        window['-OUTPUT-'].update(values['-IN-'])


window.close()