from Tix import Tk
from Tkinter import Label, Button, StringVar, Entry


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

        relevant_reference_numbers.append(DQAother_demo)
        relevant_reference_numbers.append(DQAontario_demo)
        relevant_reference_numbers.append(CQAquebec_demo)
        relevant_reference_numbers.append(CQAbc_demo)

        textLabel = Label(root, text='Entry a Refno')
        textLabel.grid(row=0, column=0)

        var = StringVar()
        enteredRefno = Entry(root, textvariable=var)
        enteredRefno.grid(row=1, column=0)

        def execute():

            # print(enteredRefno.get())
            if enteredRefno.get() == "":  # RUNNING AUTOMATIC EXECUTIONS
                if self.num == 0:
                    for i in relevant_reference_numbers:
                        if 'CQA' in i:
                            print('CQA', 'REFNO-EMPTY', i)
                        else:
                            print('DQA', 'REFNO-EMPTY', i)
                self.num+=1
            else:
                if 'CQA' in enteredRefno.get():
                    print('CQA', 'REFNO-FULL', enteredRefno.get())
                else:
                    print('DQA', 'REFNO-FULL', enteredRefno.get())
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
