from tkinter import *
from log import ExcelFile


# Designing the main menu
def main_window():
    global main_screen

    main_screen = Tk()
    main_screen.title("Finance Tracker")
    main_screen.geometry("200x150")
    Label(text = "My Finance Tracker", bg = "blue", font = ("Calibri", 13), fg = "white").pack()
    Label(text = " ").pack()

    Button(text = "Create a new workbook", command = new_workbook).pack()
    Button(text = "Open an existing workbook", command = open_existing_workbook).pack()
    Button(text = "Quit", command = main_screen.quit).pack()

    main_screen.mainloop()


# Creating a new workbook
def new_workbook():
    global new_wb_screen

    new_wb_screen = Toplevel(main_screen)  # a sub-menu of main_screen
    new_wb_screen.title("New Workbook")
    Label(new_wb_screen, text="Enter the month whose entries you want to log.").pack()
    Label(new_wb_screen, text=" ").pack()

    xl_new_workbook()  # create new xl workbook


# Retrieving new workbook title
def new_workbook_title():
    global month_title
    month_title = StringVar()
    get_month_title()
    month_entry = Entry(new_wb_screen, textvariable = month_title)
    month_entry.pack()
    Button(new_wb_screen, text="OK", command=title_and_add).pack()
    Label(new_wb_screen, text = " ").pack()


# Helper function to new_workbook_title
def title_and_add():
    get_month_title()
    xl_new.workbook_title(month_info)  # set workbook title

    add_log()  # add log to xl file


def get_month_title():
    global month_info
    month_info = month_title.get()


# Adding a log to the tracker
def add_log(screen):
    global add_log_screen

    add_log_screen = Toplevel(screen)  # sub-menu to New Workbook screen
    add_log_screen.title("Adding a log")
    log_colour = "green"

    # Labels for User entries
    Label(add_log_screen, text = "Date", fg = log_colour).grid(row = 0)
    Label(add_log_screen, text = "Location", fg = log_colour).grid(row = 1)
    Label(add_log_screen, text = "Amount", fg = log_colour).grid(row = 2)

    global date
    global location
    global amount

    date = StringVar()
    location = StringVar()
    amount = StringVar()

    Entry(add_log_screen, textvariable = date).grid(row=0, column=1)
    Entry(add_log_screen, textvariable = location).grid(row=1, column=1)
    Entry(add_log_screen, textvariable = amount).grid(row=2, column=1)

    Label(add_log_screen, text = " ").grid(row=3)

    # add user entries to xl workbook
    Button(add_log_screen, text = "OK", command = add_to_xl).grid(row=4, column=1)
    Label(add_log_screen, text = " ").grid(row=5)


# retrieves user entry
def get_date():
    global date_info
    date_info = date.get()


def get_location():
    global location_info
    location_info = location.get()


def get_amount():
    global amount_info
    amount_info = amount.get()


# activates when adding to log
def add_to_xl():
    get_date()
    get_amount()
    get_location()

    xl_new.add_log(date_info, location_info, amount_info)

    Label(add_log_screen, text = "Added data successfully").grid(row=6)
    Button(add_log_screen, text = "Done", command = close_add_log_screen).grid(row=6, column=1)
    Button(add_log_screen, text = "Add another log", command = add_log).grid(row=6, column=2)
    Label(add_log_screen, text = " ").grid(row=7)


def close_add_log_screen():
    xl_new.save_workbook()  # saves data in Excel
    add_log_screen.destroy()  # closes Date, Location, Amount window
    new_wb_screen.destroy()  # closes Enter Month window


# Create new workbook in Excel
def xl_new_workbook():
    global xl_new

    xl_new = ExcelFile()
    xl_new.new_workbook()  # create empty worksheet
    new_workbook_title()  # retrieve workbook title
    xl_new.initialise_worksheet()  # set the headers in the worksheet


# Opening an existing workbook
def open_existing_workbook():
    global open_wb_screen

    open_wb_screen = Toplevel(main_screen)  # a sub-menu of main menu
    open_wb_screen.title("Open a Workbook")
    Label(open_wb_screen, text = "Enter the name of the file to open.").pack()
    Label(open_wb_screen, text = " ").pack()

    global file_name
    file_name = StringVar()

    file_entry = Entry(open_wb_screen, textvariable = file_name)
    file_entry.pack()
    Label(open_wb_screen, text = " ").pack()

    Button(open_wb_screen, text = "OK", command = existing_workbook).pack()


def close_open_workbook_screen():
    open_wb_screen.destroy()


def get_filename():
    global file_info
    file_info = file_name.get()

# Opens existing workbook
def existing_workbook():
    xl_load = ExcelFile()  # instantiate load ExcelFile object
    get_filename()
    wb_name = str(file_info) + ".xlsx"
    print(wb_name)
    xl_load.open_workbook(wb_name)  # load workbook

    workbook_actions()

# Window to determine user action
def workbook_actions():
    action_screen = Toplevel(open_wb_screen)
    action_screen.title("Actions")
    Label(action_screen, text = "What would you like to do with this workbook?").pack()
    Button(action_screen, text = "Add log", command = lambda : add_log(open_wb_screen)).pack()
    Label(action_screen, text = " ").pack()
    Button(action_screen, text = "View workbook", command = view_wb).pack()
    Label(action_screen, text = " ").pack()

# Opens existing workbook in Excel
def view_wb():  #todo
    pass


if __name__ == "__main__":
    main_window()






