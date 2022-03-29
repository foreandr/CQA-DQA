
import SQL_CONNECTOR
import Colors
import shutil, os, sys
def DQAcoverPageWrite(CQARef, workingFolder):
    print Colors.bcolors.OKCYAN + "\nExecuting DQACover Page Write %s"%CQARef + Colors.bcolors.ENDC
    template_file = 'C:\CQA\FULL CQA - DQA\C&DQA/Templates/coverDQA.xlsx'
    saveLocation = os.path.join(r"C:\CQA\FULL CQA - DQA\C&DQA/FinishedReport-DQA", CQARef)
    dqaFile = saveLocation + r'/%sCoverDQA.xlsx' % (CQARef)

    import openpyxl
    wb = openpyxl.load_workbook(template_file)
    sheet = wb.get_sheet_by_name('CoverPage')
    sheet['B9'] = 'sdkfasdfhgasdjkhgg'
    cnx = SQL_CONNECTOR.test_connection()

    # fill in the excel sheet by checking each box that should have a number in it
    import openpyxl
    wb = openpyxl.load_workbook(template_file)
    if wb is None:
        print "Invalid Workbook"

    from openpyxl.drawing.image import Image
    os.chdir(r'C:\CQA\FULL CQA - DQA\C&DQA\Photos')
    img = Image('hs.jpg')
    sheet.add_image(img, 'B35')
    img = Image('ian.bmp')
    sheet.add_image(img, 'H35')


    wb.save(dqaFile)
    cnx.close()