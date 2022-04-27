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

import SQL_CONNECTOR
# import main
# import GUI

# GUI.Instance.execute()
# main.main()
class GUI:
    num = 0

    def main_method(self):
        root = Tk()
        root.title("CQA/DQA Generator")
        root.geometry("300x300")


        nums = Utilities.get_reference_numbers()
        relevant_reference_numbers = []

        # The automatic tests
        #for i in nums:
        #    relevant_reference_numbers.append(i)

        print()

        DQAother_demo = u'KELLY1'
        DQAontario_demo = u'KELLY'
        CQAquebec_demo = u'CQA2100409'
        CQAbc_demo = u'CQA2100540'
        CQAONt = u'CQA2200061'
        another_test = u'CQA2200124'
        CQA_ONT_FAIL = u'CQA2200094'
        N_A_TEST = u'CQA2200119'
        failtest_1 = u'CQA2200133'
        new_fail_test = U'CQA2200135'

        #relevant_reference_numbers.append(DQAother_demo)  # WORKING
        #relevant_reference_numbers.append(DQAontario_demo)  #  WORKING
        #relevant_reference_numbers.append(CQAquebec_demo) # WORKING
        #relevant_reference_numbers.append(CQAbc_demo)
        #relevant_reference_numbers.append(CQAONt)  #
        relevant_reference_numbers.append(another_test)
        #relevant_reference_numbers.append(CQA_ONT_FAIL)
        #relevant_reference_numbers.append(N_A_TEST)
        #relevant_reference_numbers.append(failtest_1)
        #relevant_reference_numbers.append(new_fail_test)

        import Colors
        print(Colors.bcolors.OKBLUE + '\n\nBEGINNING RUNNING CODE\n\n' + Colors.bcolors.ENDC)

        for i in relevant_reference_numbers:
             print(i)

        #relevant_reference_numbers = relevant_reference_numbers[0]

        textLabel = Label(root, text='Enter a Refno')
        textLabel.grid(row=0, column=0)

        var = StringVar()
        enteredRefno = Entry(root, textvariable=var)
        enteredRefno.grid(row=1, column=0)

        def execute():
            path_for_saving = r'C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport'
            if enteredRefno.get() == "":  # RUNNING AUTOMATIC EXECUTIONS
                if self.num == 0:  # so I dont run empties multiple times
                    for i in relevant_reference_numbers:
                        saveLocation = os.path.join(path_for_saving, i)
                        Utilities.makeDirectory(saveLocation)
                        if 'CQA' in i:
                            print('CQA', 'REFNO-EMPTY', i)
                            pdfConnect.pdf(str(i), saveLocation)
                            WritingPages.makeSheet(str(i), path_for_saving)
                            CoverPage.coverPageWrite(str(i), path_for_saving)
                        else:
                            print('DQA', 'REFNO-EMPTY', i)
                            DQApdfConnect.pdf(str(i), saveLocation)
                            DQAWritingPages.makeSheet_DQA(str(i), path_for_saving)
                            DQAcoverPageWrite(str(i), path_for_saving)
                self.num += 1
            else:
                saveLocation = os.path.join(path_for_saving, enteredRefno.get())
                Utilities.makeDirectory(saveLocation)
                if 'CQA' in enteredRefno.get():
                    print('CQA', 'REFNO-FULL', enteredRefno.get())
                    pdfConnect.pdf(str(enteredRefno.get()), saveLocation)
                    WritingPages.makeSheet(str(enteredRefno.get()), path_for_saving)
                    CoverPage.coverPageWrite(str(enteredRefno.get()), path_for_saving)
                else:
                    print('DQA', 'REFNO-FULL', enteredRefno.get())
                    DQApdfConnect.pdf(str(enteredRefno.get()), saveLocation)
                    DQAWritingPages.makeSheet_DQA(str(enteredRefno.get()), path_for_saving)
                    DQAcoverPageWrite(str(enteredRefno.get()), path_for_saving)

        execute()
        submitButton = Button(root, text='submit', command=execute)  # NO COMMAND
        submitButton.grid(row=2, column=0)
        root.mainloop()


Instance = GUI()
Instance.main_method()


