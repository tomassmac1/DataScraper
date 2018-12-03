import openpyxl
from tkinter import *
from tkinter import messagebox


class GUI:

    def __init__(self, master):

        self.source = str()
        self.outfile = str()
        self.master = master

        '''Creating GUI Interface'''
        frame = Frame(master, width=50, bg="white")
        frame.grid(column=0, row=0, sticky=(N, W, E, S))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        frame.pack(padx=100, pady=50)
        sourceLabel = Label(frame, text="Copy HTML Source Code Path:", bg="white", fg="black")
        sourceLabel.grid(row=1, column=0, sticky=E)
        self.HTMLSource = Entry(frame, textvariable=StringVar)
        self.HTMLSource.grid(row=1, column=1)
        destinationLabel = Label(frame, text="Copy Destination Excel Sheet Path:", bg="white", fg="black")
        destinationLabel.grid(row=2, column=0, sticky=E)
        self.destinationFile = Entry(frame, textvariable=StringVar)
        self.destinationFile.grid(row=2, column=1)
        offsetLabel = Label(frame, text="Enter row number of start of table:", bg="white", fg="black")
        offsetLabel.grid(row=3, column=0, sticky=E)
        self.offset = Entry(frame, textvariable=StringVar)
        self.offset.grid(row=3, column=1)
        colNumLabel = Label(frame, text="Enter desired column value (A=1, B=2..):", bg="white", fg="black")
        colNumLabel.grid(row=4, column=0, sticky=E)
        self.shift = Entry(frame, textvariable=StringVar)
        self.shift.grid(row=4, column=1)
        self.runButton = Label(frame, text="GO!", bg="white", fg="black", relief=FLAT)
        self.runButton.bind("<Enter>", self.hover)
        self.runButton.bind("<Leave>", self.raised)
        self.runButton.bind("<Button-1>", self.clicked)
        self.runButton.grid(row=5, column=1)

    def hover(self, event):
        self.runButton.config(relief=RAISED)

    def raised(self, event):
        self.runButton.config(relief=FLAT)

    def clicked(self, event):
        self.runButton.config(relief=SUNKEN)
        self.finish()

    def reader(self):
        '''Reads HTML text file and returns relevant data.'''
        lst = []
        lst3 = []
        self.source = self.HTMLSource.get()
        try:    # opens html file and extracts relevant <td> tags
            with open(self.source, 'r') as file:
                for line in file:
                    if '<td class ="LoggerDataCell LCF0">' in line:
                        for char in line:
                            lst.append(char)
                for char in lst:
                    if char in '0123456789':
                        lst3.append(char)
                n = 0
                while n < len(lst3): # deleting extra zero which occurs every 6th index (from <td> class)
                    del lst3[n]
                    n += 6
            return lst3
        except Exception:
            messagebox.showerror('Error!', 'No such input file or directory.')
            window.quit()

    def chunks(self, l, n):
        """Yield successive n-sized chunks from list l."""
        """Using generator reduces memory consumption"""
        for i in range(0, len(l), n):
            yield l[i:i + n]

    def finish(self):
        '''Formats extracted info and saves to workbook.'''
        final = []
        try:    # creates 6 digit strings from extracted data list.
            l_o_l = list(self.chunks(self.reader(), 6))
            for item in l_o_l:
                final.append(''.join(item))
        except Exception:
            pass
        self.outfile = self.destinationFile.get()
        try:    # input validation
            for _ in self.offset.get():
                if _ in '123456789':
                    rowNum = int(self.offset.get())
            for _ in self.shift.get():
                if _ in '123456789':
                    colNum = int(self.shift.get())
        except Exception:
            messagebox.showerror('Error!', 'Invalid row or column entry.')
            window.quit()
        try:    # creates workbook object and inputs data
            wb = openpyxl.load_workbook(self.outfile)
            sheet = wb.active
            for i in range(rowNum, len(final)+rowNum):
                c = sheet.cell(row=i, column=colNum)
                c.value = final[i-rowNum]
                wb.save(self.outfile)
            messagebox.showinfo('Success!', 'Data successfully stored.')
            window.quit()
        except Exception:
            messagebox.showerror('Error!', 'No such output file or directory.')
            window.quit()


window = Tk()
window.configure(background="white")
window.title("HTMLScraper")
w = GUI(window)
window.mainloop()
