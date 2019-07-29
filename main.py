import tkinter as tk
from xl_file import ExcelFile
from log import Log


def new_workbook():

    xl = ExcelFile()
    xl.name_worksheet()                 # name new worksheet with title of month
    xl.create_new_ws()                  # creates a new worksheet
    log = Log()
    # xl.initialise_ws()                  # create the initial format of the setup
    # wb_file_name = xl.set_wbfile_name()
    # xl.save_wb(wb_file_name)            # save the file


def main():

    window = tk.Tk()
    window.title("My Finance Logger")

    tk.Label(window, text = "My Finance Logger", fg = "white", bg = "blue").pack()     # Title

    top_frame = tk.Frame(window).pack()

    """Main menu"""

    tk.Button(top_frame, text = "Create new workbook", command = new_workbook).pack()       # Creates a new workbook and set up a new entry
    tk.Button(top_frame, text = "Open existing workbook").pack()                            # Opens an existing workbook todo
    tk.Button(top_frame, text = "Quit", command = window.quit).pack()                       # Quits the program

    """Drop-down menu"""

    root_menu = tk.Menu(window)    # root menu
    window.config(menu = root_menu)
    file_menu = tk.Menu(root_menu)
    root_menu.add_cascade(label = "File", menu = file_menu)
    file_menu.add_command(label = "Exit", command = window.quit)








    window.mainloop()



if __name__ == "__main__":
    main()







