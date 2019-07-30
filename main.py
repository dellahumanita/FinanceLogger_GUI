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


# Retrieving new workbook title
def new_workbook_title():
    global new_wb_screen

    new_wb_screen = Toplevel(main_screen)  # a sub-menu of main_screen
    new_wb_screen.title("New Workbook")
    Label(new_wb_screen, text = "Enter the month whose entries you want to log.").pack()
    Label(new_wb_screen, text = " ").pack()

    global month_title
    month_title = StringVar()

    month_entry = Entry(new_wb_screen, textvariable = month_title)
    month_entry.pack()
    Label(new_wb_screen, text = " ").pack()

    Button(new_wb_screen, text = "OK", command = get_month_title).pack()


def close_new_wb_screen():
    new_wb_screen.destroy()


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


# Adding a log to the tracker
def add_log():
    add_log_screen = Toplevel(new_wb_screen)
    add_log_screen.title("Adding a log")
    log_colour = "green"

    Label(add_log_screen, text = "Date", fg = log_colour).grid(row = 0)
    Label(add_log_screen, text = "Location", fg = log_colour).grid(row = 1)
    Label(add_log_screen, text = "Amount", fg = log_colour).grid(row = 2)

    global date_info
    global location_info
    global amount_info

    date = StringVar()
    date_info = date.get()
    location = StringVar()
    location_info = location.get()
    amount = StringVar()
    amount_info = amount.get()

    print(date, location, amount)

    Entry(add_log_screen, textvariable = date).grid(row=0, column=1)
    Entry(add_log_screen, textvariable = location).grid(row=1, column=1)
    Entry(add_log_screen, textvariable = amount).grid(row=2, column=1)

    Label(add_log_screen, text = " ").grid(row=3)

    Button(add_log_screen, text = "OK", command = xl_new_wb).grid(row=4, column=1)
    Label(add_log_screen, text = " ").grid(row=5)

def get_month_title():
    month_info = month_title.get()
    print(month_info)

def xl_new_wb():
    xl = ExcelFile()
    xl.new_wb()
    print("created new wb")
    xl.wb_title(month_info)
    print(month_info)
    print("set the file name")
    xl.initialise_ws()
    xl.add_log(date_info, location_info, amount_info)
    print("finish adding to workbook")
    xl.save_wb()
    print("end program")


if __name__ == "__main__":
    main_window()






