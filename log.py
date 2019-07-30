from openpyxl import Workbook


class ExcelFile:
    def __init__(self):
        self.title = ""
        self.wb = ""
        self.ws = ""

    def new_wb(self):
        # create a new worksheet
        self.wb = Workbook()
        self.ws = self.wb.active

    def wb_title(self, title):
        # sets the month to log as title of workbook
        self.title = title

    def open_wb(self):
        # open an existing workbook
        pass

    def initialise_ws(self):
        # setting up workbook headers
        date_c_header = self.ws['A1']
        location_c_header = self.ws['B1']
        amount_c_header = self.ws['C1']

        date_c_header.value = "DATE"
        location_c_header.value = "LOCATION"
        amount_c_header.value = "AMOUNT"

    def add_log(self, date, location, amount):
        self.ws['A2'].value = date
        self.ws['B2'].value = location
        self.ws['C2'].value = amount

    def save_wb(self):
        file_name = self.title + ".xlsx"
        self.wb.save(file_name)


# xl = ExcelFile()
# xl.new_wb()
# xl.wb_title("testfile")
# xl.initialise_ws()
# xl.save_wb()
