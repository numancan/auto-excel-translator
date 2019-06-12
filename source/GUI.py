from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext as st
from tkinter.ttk import *
from tkinter import *
from exceltranslator import Translator
import _thread

languages = [
    ('bg', 'Bulgarian'),
    ('cs', 'Czech'),
    ('zh', 'Chinese'),
    ('da', 'Danish'),
    ('de', 'German'),
    ('en', 'English'),
    ('ja', 'Japanese'),
    ('ko', 'Korean'),
    ('es', 'Spanish'),
    ('tr', 'Turkish'),
    ('de', 'German')
]

# Scale
def map(x, in_min, in_max, out_min, out_max):
    return (x - in_min) * (out_max - out_min) / (in_max - in_min) + out_min


FONT = ("Helvetica", 16)
ENTRY_WIDTH = 5

COLOR = {"line": "sea green", "ui": "dodger blue","bg":"light cyan"}


class Window():

    def __init__(self):
        self.directory = ""
        self.running = False

        # ******CREATING WINDOW*******
        self.root = Tk()
        self.root.title('Excel Translator')
        self.root.resizable(False, False)

        bg_image = PhotoImage(file="background.png")

        self.bgCanvas = Canvas(width=bg_image.width(), height=bg_image.height())
        self.bgCanvas.pack(side='top', fill='both', expand='yes')
        self.bgCanvas.create_image(0, 0, image=bg_image, anchor='nw')

        # ROW1: XLS files
        self.bgCanvas.create_text(15, 20, text="XLS Files", fill=COLOR["ui"], anchor='nw', font=FONT)
        self.bgCanvas.create_line(10, 52, 340, 52, fill=COLOR["line"])

        self.direcText = Entry(self.bgCanvas, width=22,bg=COLOR["bg"])
        self.direcText.place(x=125, y=23)

        self.browseBtn = Button(self.bgCanvas, text="Browse", width=7, bg=COLOR["ui"],
                                fg="white", command=self.askDirec)
        self.browseBtn.place(x=275, y=19)

        # ROW2: Target Language
        self.bgCanvas.create_text(14, 65, text="Target Language", fill=COLOR["ui"], anchor='nw', font=FONT)
        self.cmbTarget = Combobox(self.bgCanvas)
        self.cmbTarget['values'] = [lang[1] for lang in languages]
        self.cmbTarget.current(5)
        self.cmbTarget.place(x=185, y=67)
        self.bgCanvas.create_line(10, 97, 340, 97, fill=COLOR["line"])

        # ROW3: Column
        self.bgCanvas.create_text(14, 108, text="Translate from", fill=COLOR["ui"], anchor='nw', font=FONT)
        self.bgCanvas.create_text(198, 108, text="to", fill=COLOR["ui"], anchor='nw', font=FONT)
        self.bgCanvas.create_text(262, 108, text="to", fill=COLOR["ui"], anchor='nw', font=FONT)

        # From row
        self.fromCell = Entry(self.bgCanvas, width=ENTRY_WIDTH, justify=CENTER,bg=COLOR["bg"])
        self.fromCell.place(x=158, y=110)
        self.fromCell.insert(INSERT, "A1")

        # To row
        self.toCell = Entry(self.bgCanvas, width=ENTRY_WIDTH, justify=CENTER,bg=COLOR["bg"])
        self.toCell.place(x=222, y=110)
        self.toCell.insert(INSERT, "END")

        # To column
        self.toColumn = Entry(self.bgCanvas, width=ENTRY_WIDTH, justify=CENTER,bg=COLOR["bg"])
        self.toColumn.place(x=286, y=110)
        self.toColumn.insert(INSERT, "B")

        self.bgCanvas.create_line(10, 140, 340, 140, fill=COLOR["line"])

        # ROW4: Progress Bar
        self.s = Style(self.bgCanvas)
        self.s.layout("LabeledProgressbar",
                      [('LabeledProgressbar.trough',
                        {'children': [('LabeledProgressbar.pbar',
                                       {'side': 'left', 'sticky': 'ns'}),
                                      ("LabeledProgressbar.label",
                                       {"sticky": ""})],
                         'sticky': 'nswe'})])
        self.progBar = Progressbar(self.bgCanvas, orient="horizontal", length=318, style="LabeledProgressbar")
        self.s.configure("LabeledProgressbar", text="0 %      ", foreground='black', background=COLOR["ui"])
        self.progBar.place(x=15, y=155)

        # ROW5: Start Button
        self.startBtn = Button(self.bgCanvas, text="Start", width=10, height=2, bg=COLOR["ui"], fg="white", command=self.start)
        self.startBtn.place(x=140, y=185)

        self.root.mainloop()

    def message(self, typ, text):
        if typ == "error":
            messagebox.showerror("Error", text)
        elif typ == "info":
            messagebox.showinfo("Information", text)
        elif typ == "retry":
            return messagebox.askretrycancel("Error", text)

    def askDirec(self):
        self.directory = filedialog.askopenfilenames()
        if self.directory == "":
            self.message("error", "XLS files not selected!")
        else:
            self.changeEntry(self.direcText, " {0} XML selected!".format(len(self.directory)))

    def changeEntry(self, entry, text):
        entry.delete(0, END)
        entry.insert(INSERT, text)
        entry.update()

    def updateProgBar(self, current,minRow,maxRow, finished, xlsname):
        progVal = round(map(current, minRow, maxRow, 0, 100), 2)
        self.progBar["value"] = progVal
        self.s.configure("LabeledProgressbar", text="[{0}/{1}] {3} {2}%".format(finished,
                                                                                len(self.directory), progVal, xlsname))
        self.progBar.update()

    def getInputs(self):
        lang=[lang[0] for lang in languages][[lang[1] for lang in languages].index(self.cmbTarget.get())]
        return lang, self.fromCell.get(),self.toCell.get(), self.toColumn.get()

    def restart(self):
        self.direcText.delete(0, END)
        self.directory = None
        self.progBar["value"] = 0
        self.s.configure("LabeledProgressbar", text="0 %      ")
        self.progBar.update()
        self.startBtn["state"] = NORMAL
        self.startBtn.configure(bg=COLOR["ui"])

    def start(self):
        if self.directory == "" or self.directory == None:
            self.message("error", "Please select XLS files!")
        else:
            _thread.start_new_thread(Translator, (self,))
            self.startBtn["state"] = DISABLED
            self.startBtn.configure(bg="gray")
