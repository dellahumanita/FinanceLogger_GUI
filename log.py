import tkinter as tk


class Log:  #todo

    def __init__(self):
        self.window = tk.Tk()
        self.log_colour = "green"
        tk.Label(self.window, text="Entry Actions").pack()
        self.add_btn = tk.Button(self.window, text = "Add Entry", fg = self.log_colour, command = self.add)
        self.view_btn = tk.Button(self.window, text = "View entries", fg = self.log_colour)
        self.close_btn = tk.Button(self.window, text = "Close", fg = self.log_colour, command = self.window.quit)
        self.add_btn.pack()
        self.view_btn.pack()
        self.close_btn.pack()

    def add(self):  #todo
        # GUI
        add_window = tk.Tk()
        add_window.title("Adding an entry")
        tk.Label(add_window, text = "Entry number").grid(row = 0)
        tk.Entry(add_window).grid(row = 0, column = 1)
        tk.Label(add_window, text = "Date (DD-MM-YYYY)").grid(row = 1, column = 0)
        tk.Entry(add_window).grid(row = 1, column = 1)
        tk.Label(add_window, text = "Location").grid(row = 2)
        tk.Entry(add_window).grid(row = 2, column = 1)
        tk.Label(add_window, text = "Amount spent").grid(row = 3)
        tk.Entry(add_window).grid(row = 3, column = 1)

    def view(self):  #todo
        # opens the excel file and display entries
        pass

    def edit(self):  #todo
        # selects an entry and edits the data
        pass

    def save(self):  #todo
        # saves the data
        pass


