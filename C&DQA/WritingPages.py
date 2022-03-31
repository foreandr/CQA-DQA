import openpyxl
import Colors
from Utilities import findLocation
import NONOntarioCQAReport

def makeSheet(CQARef, workingFolder):
    '''Runs the printing function depending on the location'''

    # Gets the location using the findLocation function
    location = findLocation(CQARef)
    print CQARef + " " + Colors.bcolors.OKGREEN + "Current Location: " + location + Colors.bcolors.ENDC
    import OntarioCQAReport
    if location == 'ON' or location == 'QC':
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\TEMPLATE ON WRITTEN.xlsx'
        wb = openpyxl.load_workbook(templateFile)
        OntarioCQAReport.OntarioQuebecCQA(wb, CQARef)
    else:
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\report.xlsx'
        wb = openpyxl.load_workbook(templateFile)
        NONOntarioCQAReport.BCandOtherReport(wb, CQARef)
