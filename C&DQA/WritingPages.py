import openpyxl

import Colors
from Utilities import findLocation
from OntarioPart import findOntarioCatagory, OntarioResults, OntarioPrint
from CCMEPart import findCCMECatagory, CCMEResults, CCMEPrint
from QuebecPart import findQuebecCatagory, QuebecResults, QuebecPrint
from BCPart import findBCCatagory, BCResults, BCPrint
import NONOntarioCQAReport
def makeSheet(CQARef, workingFolder):
    '''Runs the printing function depending on the location'''
    
    #Gets the location using the findLocation function
    location = findLocation(CQARef)
    print Colors.bcolors.OKGREEN + "Current Location: " + location + Colors.bcolors.ENDC
    import OntarioCQAReport
    if location == 'ON' or location == 'QC':
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\TEMPLATE ON WRITTEN.xlsx'
        wb = openpyxl.load_workbook(templateFile)
        OntarioCQAReport.OntarioQuebecCQA(wb, CQARef)
    else:
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\TEMPLATE BC WRITTEN.xlsx'
        wb = openpyxl.load_workbook(templateFile)
        NONOntarioCQAReport.BCandOtherReport(wb, CQARef)

    
