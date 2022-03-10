import os
import shutil
import DQAWritingPages
import SQL_CONNECTOR
import Utilities
from WritingPages import makeSheet
import ttk, os
from Tkinter import *
import sys


def main():
    root = Tk()
    root.title("DQA Generator")
    root.geometry("200x75")

    other_demo = u'KELLY1' # PEI
    ontario_demo = u'KELLY'

    DQATestList = []
    DQATestList.append(other_demo)
    DQATestList.append(ontario_demo)

    def rungen():
        path_for_saving = r'C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport'
        saveLocation = os.path.join(path_for_saving, enteredRefno.get())
        Utilities.makeDirectory(saveLocation)

        DQAWritingPages.makeSheet_DQA(enteredRefno.get(), path_for_saving)
        import DQApdfConnect
        DQApdfConnect.pdf(enteredRefno.get(), saveLocation)

    textLabel = Label(root, text='Entry a Refno')
    textLabel.grid(row=0, column=0)

    var = StringVar()
    enteredRefno = Entry(root, textvariable=var)
    enteredRefno.grid(row=1, column=0)

    enteredRefno = Entry(root, text="Enter Refno:")
    enteredRefno.grid(row=1, column=0)
    import pdfConnect
    for i in DQATestList:
        enteredRefno.insert(END, i)
        rungen()
        enteredRefno.delete(0, END)

    submitButton = Button(root, text='submit', command=rungen)
    submitButton.grid(row=2, column=0)

    root.mainloop()

if __name__ == "__main__":
    main()