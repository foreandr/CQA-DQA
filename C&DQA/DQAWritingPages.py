import openpyxl
from Utilities import findLocation
import OntarioDQAW
import CFIADQA


def makeSheet_DQA(CQARef, workingFolder):
    location = findLocation(CQARef)
    if location == 'ON':
        print ("Location: %s" % location)
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\Ontario DQA -W.xlsx'
        wb = openpyxl.load_workbook(templateFile)
        OntarioDQAW.OntarioPrintDQA(wb, CQARef)
    else:
        print ("Location: %s" % location)
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\CFIA DQA-KO.xlsx'
        wb = openpyxl.load_workbook(templateFile)
        # print ("Location: %s" % location)
        CFIADQA.CFIAPrintDQA(wb, CQARef)
