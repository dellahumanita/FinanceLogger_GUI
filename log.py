from openpyxl import Workbook, load_workbook


class ExcelFile:
    def __init__(self):
        self.title = ""
        self.wb = None  # new workbook
        self.ws = None  # new worksheet
        self.load_wb = None  # existing workbook

    def new_wb(self):
        # create a new worksheet
        self.wb = Workbook()
        self.ws = self.wb.active

    def wb_title(self, title):
        # sets the month to log as title of workbook
        self.title = title

    def open_wb(self, filename):
        # open an existing workbook
        self.load_wb = load_workbook(filename)
        print("wb opened")

    def initialise_ws(self):
        # setting up workbook headers

        date_c_header = self.ws['A1']
        location_c_header = self.ws['B1']
        amount_c_header = self.ws['C1']

        date_c_header.value = "DATE"
        location_c_header.value = "LOCATION"
        amount_c_header.value = "AMOUNT"

    def add_log(self, date, location, amount):
        # adds log to the following row

        last_row = self.find_last_row_used()
        next_row = last_row+1
        self.ws.cell(column=1, row=next_row, value=date)
        self.ws.cell(column=2, row=next_row, value=location)
        self.ws.cell(column=3, row=next_row, value=amount)

    def find_last_row_used(self):
        # searches for the last row entered
        return self.ws.max_row

    def save_wb(self):
        file_name = self.title + ".xlsx"
        self.wb.save(file_name)


# xl = ExcelFile()
# xl.new_wb()
# xl.wb_title("testfile")
# xl.initialise_ws()
# xl.add_log('01/02/2019','No Frills','30')
# xl.add_log('02/03/2019','no','10')
# xl.add_log('03/04/2019', 'Chapters', '100')
# #xl.add_log_v2()
# xl.save_wb()

