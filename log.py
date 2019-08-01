from openpyxl import Workbook


class ExcelFile:
    def __init__(self):
        self.title = ""
        self.wb = None
        self.ws = None

    def new_wb(self):
        # create a new worksheet
        self.wb = Workbook()
        self.ws = self.wb.active

    def wb_title(self, title):
        # sets the month to log as title of workbook
        self.title = title

    def open_wb(self):  #todo
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
        for row in self.ws.iter_rows(min_row=2, max_col=3, max_row=5):  # iterate over rows 2-100
            for cell in row:  # for each cell in the row
                if cell.value is None:  # if the cell is empty
                    if cell.column == 1:
                        cell.value = date
                    elif cell.column == 2:
                        cell.value = location
                    elif cell.column == 3:
                        cell.value = float(amount)
                else:  # iterate to the next row
                    continue
            break

    def save_wb(self):
        file_name = self.title + ".xlsx"
        self.wb.save(file_name)

def testing():
    xl = ExcelFile()
    xl.new_wb()
    xl.wb_title("testfile")
    xl.initialise_ws()
    xl.add_log('01/02/2019','No Frills','30')
    xl.add_log('02/02/2019', "Dollarama", "10")
    xl.save_wb()
    
testing()

