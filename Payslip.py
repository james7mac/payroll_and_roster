import tabula, os, csv, datetime

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
