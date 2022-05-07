import sys
import time
from tkinter import Tk, Button, Frame, Label, StringVar, Entry
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename, asksaveasfilename
from BackBone import *
from tkinter import INSERT
import win32gui, win32con
the_program_to_hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(the_program_to_hide , win32con.SW_HIDE)
Year = "2021/2022"
Term = "Mid-Term Second Semister"

class PrintLogger(object):  # create file like object

    def __init__(self, textbox):  # pass reference to text widget
        self.textbox = textbox  # keep ref

    def write(self, text):
        self.textbox.configure(state="normal")  # make field editable
        self.textbox.insert("end", text)  # write text to textbox
        self.textbox.see("end")  # scroll to end
        self.textbox.configure(state="disabled")  # make field readonly

    def flush(self):  # needed for file like object
        pass


class MainGUI(Tk ):

    def __init__(self , className):
        self.startP = False
        self.SOURCE_DIRECTORY = os.getcwd()
        self.FONT = ("consolas", "20", "normal")
        self.FONT2 = ("consolas", "16", "normal")
        self.FONT3 = ("consolas", "10", "normal")
        Tk.__init__(self , className=className)
        self.geometry('700x450')
        self.root = Frame(self )
        self.root.pack()
        self.log_widget = ScrolledText(self.root, height=9, width=80, font=("consolas", "8", "normal"))

        self.L_Greating = Label(self.root, text="Welcome to BIS Certificate Maker" , font=self.FONT )
        self.L_Greating.grid(row=0, column=0,padx=5 ,pady=35 , columnspan=4)
        self.L_Greating = Label(self.root, text="Please Select Your Files To Begin" , font=self.FONT2)
        self.L_Greating.grid(row=1, column=0 ,padx=5 , pady=5 , columnspan=4)
        self.L_choseFile = Label(self.root , text= "Chose Excel File " , font=self.FONT2)
        self.L_choseFile.grid(row = 2, column = 0, pady = 30 )
        self.browse_b = StringVar()
        self.browse_b.set('Browse')
        self.B_Browse_Excel = Button(self.root, command=self.choseExcel, textvariable= self.browse_b, font=self.FONT3, height=1, width=25)
        self.B_Browse_Excel.grid(columnspan=3, column=2, row=2)

        self.L_choseFile2 = Label(self.root, text="Chose Certificate File ", font=self.FONT2)
        self.L_choseFile2.grid(row=3, column=0, pady=0)
        self.browse_b2 = StringVar()
        self.browse_b2.set('Browse')
        self.B_Browse_word = Button(self.root, command=self.choseWord, textvariable=self.browse_b2, font=self.FONT3,height=1, width=25)
        self.B_Browse_word.grid(columnspan=3, column=2, row=3)

        self.L_choseClass = Label(self.root, text="Chose Class ", font=self.FONT2)
        self.L_choseClass.grid(row=4, column=0, pady=20)
        self.Slas = Entry (self.root)
        self.Slas.grid(columnspan=3,row=4, column=1, pady=20)
        self.startbut = StringVar()
        self.startbut.set('Start')
        self.B_Start = Button(self.root,  command=self.start ,textvariable= self.startbut, font=self.FONT2,
                                    height=1, width=25)
        self.B_Start.grid(columnspan=4, column=0, row=5)

    def choseExcel (self) :
        self.excelFile = askopenfilename(initialdir=self.SOURCE_DIRECTORY, title="Select Excel", filetypes=([("Excel files", ".xls .xlsx")]))
        self.exFN = self.excelFile.split('/')
        self.exFN = self.exFN[-1]
        self.browse_b.set (self.exFN)
    def choseWord (self) :
        self.wordFile = askopenfilename(initialdir=self.SOURCE_DIRECTORY, title="Select Word", filetypes=([("Word files", ".doc .docx")]))
        self.wdFN = self.wordFile.split('/')
        self.wdFN = self.wdFN[-1]
        self.browse_b2.set (self.wdFN)

    def start (self) :
        self.startbut.set('Pleas wait..')
        self.startP = True
        self.sclass = self.Slas.get()
        self.printer(self.startP)
        self.prog()
        self.startbut.set('Start')

    def prog (self):
        global the_program_to_hide
        win32gui.ShowWindow(the_program_to_hide, win32con.SW_SHOW)
        self.startbut.set("getting data from excel...")
        self.studentsdata = getDataFromExcel(self.excelFile)
        self.startbut.set("Setting Ranks..")
        self.studentsdata = rankStudents(self.studentsdata)
        self.startbut.set("Filling Files..")
        self.files = fillWord(self.wordFile, self.studentsdata, Year, Term, self.sclass)
        self.startbut.set('converting to PDF..')
        convert2PDF(self.sclass)
        self.startbut.set('Merging Files')
        finalFile = mergeFiles(self.sclass, self.files)
        print("done , openning file..")
        time.sleep(5)
        the_program_to_hide = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(the_program_to_hide, win32con.SW_HIDE)
        os.startfile(finalFile)

    def printer (self,printing) :
        self.log_widget.insert(INSERT, f'{printing}\n')

