import openpyxl

import Colors
from Utilities import findLocation
from OntarioPart import findOntarioCatagory, OntarioResults, OntarioPrint
from CCMEPart import findCCMECatagory, CCMEResults, CCMEPrint
from QuebecPart import findQuebecCatagory, QuebecResults, QuebecPrint
from BCPart import findBCCatagory, BCResults, BCPrint

#using location found in utilities the correct report template is found and used Coverpage Category can change depending on provice for pass or fail rules
# coverpage is printed first and then the report page. Both DO depend on province
def findCoverCatagory(CQARef):
    '''finds the catagory for the cover page by running a function relating to its location'''
    location = findLocation(CQARef)
    if location == "ON":
        return findOntarioCatagory(CQARef)
    elif location == 'QC':
        return findQuebecCatagory(CQARef)
    else:
        return findCCMECatagory(CQARef)
    
def makeSheet(CQARef, workingFolder):
    '''Runs the printing function depending on the location'''
    
    #Gets the location using the findLocation function
    location = findLocation(CQARef)
    print Colors.bcolors.OKGREEN + "Current Location: " + location + Colors.bcolors.ENDC
    
    finalResult = None
    #This basically gets the results as a dictonary 
    #If its ontario then get the ontario results and set the template file
    
    #!!!!Changed for testing Quebec formats (NEED TO ADD BC TEMPLATE TO CODE)
    if location == "ON":
        finalResult = OntarioResults(CQARef)
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\TEMPLATE ON WRITTEN.xlsx'
    
    elif location == "QC":
        finalResult = QuebecResults(CQARef)
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\TEMPLATE ON WRITTEN.xlsx' # CHANGED FROM QC TO ON
    
    #If its anything else then get the ccme results and set the template file
    else:
        finalResult = CCMEResults(CQARef)
        templateFile = r'C:\CQA\FULL CQA - DQA\C&DQA\Templates\report.xlsx'
    
    #opens the workbook
    wb = openpyxl.load_workbook(templateFile)
    if wb is None:
        print Colors.bcolors.FAIL + "Invalid Workbook" + Colors.bcolors.ENDC
    
    #this uses the dictionary to write to results to the excel page
    #Calls function to write to the excel file depending on the province
    if location == 'ON':
        OntarioPrint(wb, finalResult, CQARef)
    elif location == 'QC':
        OntarioPrint(wb, finalResult, CQARef) # switched from quebec print to ontario print
    else:
        CCMEPrint(wb, finalResult, CQARef)
    
