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

    xl_new_wb()  # create new xl workbook


# Retrieving new workbook title
def new_workbook_title():
    global month_title
    month_title = StringVar()

    month_entry = Entry(new_wb_screen, textvariable = month_title)
    month_entry.pack()
    Button(new_wb_screen, text="OK", command=title_and_add).pack()
    Label(new_wb_screen, text = " ").pack()
    
    get_month_title()
    print(month_info, "is the month info")
    xl.wb_title(month_info)  # set workbook title
    

# Helper function to new_workbook_title
def title_and_add():
    get_month_title()

    add_log()  # add log to xl file


def get_month_title():
    global month_info
    month_info = month_title.get()


# Adding a log to the tracker
def add_log():
    global add_log_screen

    add_log_screen = Toplevel(new_wb_screen)  # sub-menu to New Workbook screen
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

    xl.add_log(date_info, location_info, amount_info)
    
    Label(add_log_screen, text = "Added data successfully").grid(row=6)
    Button(add_log_screen, text = "Done", command = close_add_log_screen).grid(row=6, column=1)
    Button(add_log_screen, text = "Add another log", command = add_log).grid(row=6, column=2)
    Label(add_log_screen, text = " ").grid(row=7)

def close_add_log_screen():
    xl.save_wb()
    add_log_screen.destroy()


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

    Button(open_wb_screen, text = "OK", command = close_new_wb_screen).pack()


def close_open_wb_screen():
    open_wb_screen.destroy()


def xl_new_wb():
    global xl

    xl = ExcelFile()
    xl.new_wb()  # create empty worksheet
    new_workbook_title()  # retrieve workbook title
    xl.initialise_ws()  # set the headers in the worksheet


if __name__ == "__main__":
    main_window()




# testing new change

