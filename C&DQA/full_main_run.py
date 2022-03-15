import os
from Tix import Tk
from Tkinter import Label, Button, StringVar, Entry
import Utilities
import DQAWritingPages
import DQApdfConnect
from CoverPageDQA import DQAcoverPageWrite
import CoverPage
import WritingPages
import pdfConnect
import OntarioCQAReport
# import main
# import GUI

# GUI.Instance.execute()
# main.main()
class GUI:
    num = 0
    def main_method(self):
        root = Tk()
        root.title("CqA/DQA Generator")
        root.geometry("300x300")

        relevant_reference_numbers = []
        DQAother_demo = u'KELLY1'  # PEI
        DQAontario_demo = u'KELLY'
        CQAquebec_demo = u'CQA2100409'  # quebec tester
        CQAbc_demo = u'CQA2100540'  # bc tester

        # current_test = u'CQA2200061'  #MAKE OR BREAKs

        #relevant_reference_numbers.append(DQAother_demo)
        #relevant_reference_numbers.append(DQAontario_demo)
        relevant_reference_numbers.append(CQAquebec_demo)
        #relevant_reference_numbers.append(CQAbc_demo)
        #relevant_reference_numbers.append(current_test)

        textLabel = Label(root, text='Entry a Refno')
        textLabel.grid(row=0, column=0)

        var = StringVar()
        enteredRefno = Entry(root, textvariable=var)
        enteredRefno.grid(row=1, column=0)

        def execute():
            path_for_saving = r'C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport'
            if enteredRefno.get() == "":  # RUNNING AUTOMATIC EXECUTIONS
                if self.num == 0:  #so I dont run empties multiple times
                    for i in relevant_reference_numbers:
                        saveLocation = os.path.join(path_for_saving, i )
                        Utilities.makeDirectory(saveLocation)
                        if 'CQA' in i:
                            print('CQA', 'REFNO-EMPTY', i)
                            CoverPage.coverPageWrite(str(i), path_for_saving)
                            WritingPages.makeSheet(str(i), path_for_saving)
                            pdfConnect.pdf(str(i), saveLocation)
                        else:
                            print('DQA', 'REFNO-EMPTY', i)
                            DQAcoverPageWrite(str(i), path_for_saving)
                            DQAWritingPages.makeSheet_DQA(str(i), path_for_saving)
                            DQApdfConnect.pdf(str(i), saveLocation)
                self.num+=1
            else:
                saveLocation = os.path.join(path_for_saving, enteredRefno.get())
                Utilities.makeDirectory(saveLocation)
                if 'CQA' in enteredRefno.get():
                    print('CQA', 'REFNO-FULL', enteredRefno.get())
                    CoverPage.coverPageWrite(str(enteredRefno.get()), path_for_saving)
                    WritingPages.makeSheet(str(enteredRefno.get()), path_for_saving)
                    pdfConnect.pdf(str(enteredRefno.get()), saveLocation)
                else:
                    print('DQA', 'REFNO-FULL', enteredRefno.get())
                    DQAcoverPageWrite(str(enteredRefno.get()), path_for_saving)
                    DQAWritingPages.makeSheet_DQA(str(enteredRefno.get()), path_for_saving)
                    DQApdfConnect.pdf(str(enteredRefno.get()), saveLocation)

        execute()
        submitButton = Button(root, text='submit', command=execute)  # NO COMMAND
        submitButton.grid(row=2, column=0)

        root.mainloop()

Instance = GUI()
Instance.main_method()

'''
SHIT TO DO TOMORROW:
TITLE PAGE
COMPLETE UNIFICATION
SOMETHING ABOUT NEGATIVES

'''
