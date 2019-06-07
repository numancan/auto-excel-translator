from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext as st
from tkinter.ttk import *
from tkinter import *
from exceltranslator import Translator
import _thread

LANGUAGES = ("English", "Turkce",)

# Scaling for progress bar


def map(x, in_min, in_max, out_min, out_max):
    return (x - in_min) * (out_max - out_min) / (in_max - in_min) + out_min


FONT = ("Helvetica", 16)
ENTRY_WIDTH = 5
COLOR = "dodger blue"


class Window():

    def __init__(self):
        self.directory = ""
        self.running = False

        # ******CREATING WINDOW*******
        self.root = Tk()
        self.root.title('Excel Translater')
        self.root.resizable(False, False)

        bg_image = PhotoImage(file="background.png")

        self.bgCanvas = Canvas(width=bg_image.width(), height=bg_image.height())
        self.bgCanvas.pack(side='top', fill='both', expand='yes')
        self.bgCanvas.create_image(0, 0, image=bg_image, anchor='nw')

        # ROW1: XLS files
        self.bgCanvas.create_text(15, 20, text="XLS Files", fill=COLOR, anchor='nw', font=FONT)
        self.bgCanvas.create_line(10, 52, 340, 52, fill="white")

        self.direcText = Entry(self.bgCanvas, width=22)
        self.direcText.place(x=125, y=23)

        self.browseBtn = Button(self.bgCanvas, text="Browse", width=7, bg=COLOR,
                                fg="white", command=self.askDirec)
        self.browseBtn.place(x=275, y=19)

        # ROW2: Target Language
        self.bgCanvas.create_text(14, 65, text="Target Language", fill=COLOR, anchor='nw', font=FONT)
        self.cmbTarget = Combobox(self.bgCanvas)
        self.cmbTarget['values'] = ("English", "Turkce")
        self.cmbTarget.current(0)
        self.cmbTarget.place(x=185, y=67)
        self.bgCanvas.create_line(10, 97, 340, 97, fill="white")

        # ROW3: Column
        self.bgCanvas.create_text(14, 105, text="Column:", fill=COLOR, anchor='nw', font=FONT)
        self.bgCanvas.create_text(30, 130, text="Source", fill=COLOR, anchor='nw', font=("Helvetica", 14))
        self.bgCanvas.create_text(200, 130, text="Target", fill=COLOR, anchor='nw', font=("Helvetica", 14))

        self.srcCol = Entry(self.bgCanvas, width=ENTRY_WIDTH)
        self.srcCol.place(x=100, y=134)
        self.srcCol.insert(INSERT, 0)

        self.trgtCol = Entry(self.bgCanvas, width=ENTRY_WIDTH)
        self.trgtCol.place(x=265, y=134)
        self.trgtCol.insert(INSERT, 1)

        self.bgCanvas.create_line(10, 163, 340, 163, fill="white")

        # ROW4: Starting rows
        self.bgCanvas.create_text(14, 170, text="Starting Row:", fill=COLOR, anchor='nw', font=FONT)
        self.bgCanvas.create_text(30, 195, text="Source", fill=COLOR, anchor='nw', font=("Helvetica", 14))
        self.bgCanvas.create_text(200, 195, text="Target", fill=COLOR, anchor='nw', font=("Helvetica", 14))

        self.strtTrgt = Entry(self.bgCanvas, width=ENTRY_WIDTH)
        self.strtTrgt.place(x=100, y=198)
        self.strtTrgt.insert(INSERT, 6)

        self.strtSrc = Entry(self.bgCanvas, width=ENTRY_WIDTH)
        self.strtSrc.place(x=265, y=198)
        self.strtSrc.insert(INSERT, 6)

        self.bgCanvas.create_line(10, 226, 340, 226, fill="white")

        # ROW5: Progress Bar
        self.s = Style(self.bgCanvas)
        self.s.layout("LabeledProgressbar",
                      [('LabeledProgressbar.trough',
                        {'children': [('LabeledProgressbar.pbar',
                                       {'side': 'left', 'sticky': 'ns'}),
                                      ("LabeledProgressbar.label",
                                       {"sticky": ""})],
                         'sticky': 'nswe'})])
        self.progBar = Progressbar(self.bgCanvas, orient="horizontal", length=318, style="LabeledProgressbar")
        self.s.configure("LabeledProgressbar", text="0 %      ", foreground='black', background=COLOR)
        self.progBar.place(x=15, y=238)

        # ROW6: Start Button
        self.startBtn = Button(self.bgCanvas, text="Start", width=10, height=2, bg=COLOR, fg="white", command=self.start)
        self.startBtn.place(x=140, y=270)

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

    def updateProgBar(self, current, maxRow, finished, xlsname):
        progVal = round(map(current, int(self.strtSrc.get()) - 1, maxRow - 1, 0, 100), 2)
        self.progBar["value"] = progVal
        self.s.configure("LabeledProgressbar", text="[{0}/{1}] {3} {2}%".format(finished,
                                                                                len(self.directory), progVal, xlsname))
        self.progBar.update()

    def getInputs(self):
        return self.cmbTarget.get(), int(self.srcCol.get()), int(self.trgtCol.get()), int(self.strtSrc.get()), int(self.strtTrgt.get())

    def restart(self):
        self.direcText.delete(0, END)
        self.directory = None
        self.progBar["value"] = 0
        self.s.configure("LabeledProgressbar", text="0 %      ")
        self.progBar.update()
        self.startBtn["state"] = NORMAL
        self.startBtn.configure(bg=COLOR)

    def start(self):
        if self.directory == "" or self.directory == None:
            self.message("error", "Please select XLS files!")
        else:
            _thread.start_new_thread(Translator, (self,))
            self.startBtn["state"] = DISABLED
            self.startBtn.configure(bg="gray")
