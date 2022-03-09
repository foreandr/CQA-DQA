import collections
import tkSimpleDialog
import ttk, os
from Tkinter import *
from CoverPage import coverPageWrite
from WritingPages import makeSheet
import pdfConnect
import shutil
import Colors
import SQL_CONNECTOR
import pynput.keyboard


class Gui:
    '''
    Purpose:
    Date:
    '''

    def __init__(self):
        pass

    '''-----Important Data-----'''
    connection = SQL_CONNECTOR.test_connection()
    cQaTestList = []  # this will be being produced by kelly's queries in the future
    relevant_reference_numbers = []

    quebec_demo = u'CQA2100409'  # quebec tester
    bc_demo = u'CQA2100540'  # bc tester
    ontario_demo = u'CQA2200021'
    ontario_demo_errors = u'CQA2100100'

    #relevant_reference_numbers.append(quebec_demo)
    # relevant_reference_numbers.append(bc_demo)
    relevant_reference_numbers.append(ontario_demo)
    #relevant_reference_numbers.append(ontario_demo_errors)

    for i in relevant_reference_numbers:
        cQaTestList.append(i)
    '''-----Functions-----'''

    def caps(self, event):
        var.set(var.get().upper())

    def runGen(self):
        '''
        Purpose:
        Date:
        Returns
        '''
        print Colors.bcolors.OKCYAN + "\nExecuting RunGEN/GUI" + Colors.bcolors.ENDC
        # COULD BE ANY PATH HERE WANTED
        path_for_saving = r'C:\CQA\FULL CQA - DQA\C&DQA\FinishedReport'
        # path_for_saving = r'C:\CQA\NewFormatCQA\FinishedReport'
        global entryWidget
        global progress
        print Colors.bcolors.OKGREEN + entryWidget.get() + Colors.bcolors.ENDC

        saveLocation = os.path.join(path_for_saving, str(entryWidget.get()))
        try:
            os.mkdir(saveLocation)
        except:
            shutil.rmtree(saveLocation)
            os.mkdir(saveLocation)

        progress.start()
        # These call the writing methods. which write the file and save it to the location specified
        coverPageWrite(entryWidget.get(), path_for_saving)
        makeSheet(entryWidget.get(), path_for_saving)
        pdfConnect.pdf(entryWidget.get(), saveLocation)

    def execute(self):
        '''
        Purpose:
        Date:
        Returns
        '''

        list = self.get_reference_numbers(self.connection)

        for i in list:
            self.cQaTestList.append(i)

        global var
        global entryWidget
        global progress

        root = Tk()
        root.title("CQA Generator")
        root["padx"] = 40
        root["pady"] = 20

        # Create a text frame to hold the text Label and the Entry widget
        textFrame = Frame(root)
        progress = ttk.Progressbar(root, length='500')

        # Create a Label in textFrame
        entryLabel = Label(textFrame)
        entryLabel["text"] = "Enter the CQA Number:"
        entryLabel.pack(side=TOP)

        # Create an Entry Widget in textFrame
        var = StringVar()
        entryWidget = Entry(textFrame, textvariable=var)
        for i in self.cQaTestList:
            entryWidget.insert(END, i)
            Instance.runGen()
            entryWidget.delete(0, END)

        entryWidget["width"] = 50
        entryWidget.pack(side=LEFT)
        entryWidget.bind("<KeyRelease>", self.caps)

        textFrame.pack()
        button = Button(root, text="Generate CQA", command=self.runGen)
        button.pack()
        root.mainloop()

    def get_reference_numbers(self, connection):
        print Colors.bcolors.UNDERLINE + Colors.bcolors.HEADER + "Executing Kelly's Queries" + Colors.bcolors.ENDC
        reference_number_index = 6
        cursor = connection.cursor()
        query = '''
        SELECT report.rpt_name,

               report.custno,

               report.module,

               report.rptno,

               report.company,

               report.grow_1,

               report.refno,

               report.rpt_status,

               report.create_date,

               report.state

          FROM alms.report report

        WHERE (report.rpt_name = 'SQA_COMP'

        OR report.rpt_name ='STP'

        OR report.rpt_name = 'AL_STP'

        OR report.rpt_name='AL_CQA-O'

        OR report.rpt_name = "AL-ON-CQ"

        OR report.rpt_name = "AL_CQA")

        AND (rpt_status <="6")
        '''
        cursor.execute(query)
        a_dict = collections.defaultdict(list)
        for i in cursor:
            if i[reference_number_index] != '':  # if it has a reference number
                # print i[6]
                a_dict[i[reference_number_index]].append(i)
                pass

        '''
        With the reference number as the key, check each list associated with the ref number for values above 5
        '''
        usable_ref_numbers = []
        temp_list = []
        for key, value in a_dict.iteritems():
            for i in value:
                # print key + ": " + str(i[7])
                temp_list.append(i[7])

            # check the list for the numbers being over 5
            if 5 <= temp_list[0] < 7 and 5 <= temp_list[1] < 7:
                print key + " can be used " + str(temp_list)
                usable_ref_numbers.append(key)
            else:
                print key + " cannot be used " + str(temp_list)
                pass
            # reset the list
            temp_list = []
        cursor.close()
        print Colors.bcolors.HEADER + "---------------------------------" + Colors.bcolors.ENDC
        return usable_ref_numbers



Instance = Gui()

