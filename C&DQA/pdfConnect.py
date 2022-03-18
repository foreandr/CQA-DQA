import os, sys, shutil
import mysql.connector
from mysql.connector import errorcode
import pdfDownload
import Colors
import SQL_CONNECTOR

def pdf(CQARef, outputFolder):
    print Colors.bcolors.OKCYAN + "\nExecuting PDF CQA Page Write" + Colors.bcolors.ENDC
    # Connects to sql server and if it doesnt work returns an error
    cnx = SQL_CONNECTOR.test_connection()

    cursor = cnx.cursor()
    query = "SELECT COUNT(*) FROM alms.report WHERE refno='%s'" % (CQARef)
    cursor.execute(query)
    for item in cursor:
        reportNumCount = int(item[0])

    print Colors.bcolors.OKGREEN + r"Number of Report is %d" % reportNumCount + Colors.bcolors.ENDC

    # Query out the report numbers
    cursor = cnx.cursor()
    query = "SELECT * FROM alms.report WHERE refno='%s'" % (CQARef)
    cursor.execute(query)
    reportNumbers = []
    for item in cursor:
        print Colors.bcolors.HEADER + "[Ontario Report REF Query:]" + Colors.bcolors.ENDC + str(item)
        tempList = [item[0], item[1]]  # get report # and report type
        reportNumbers.append(tempList)

    # Store the report numbers
    if reportNumCount == 2:
        soilReport = None
        envReport = None
        for item in reportNumbers:
            if str(item[1]) == "SOIL":
                soilReport = str(item[0])
            elif str(item[1]) == "ENVI":
                #print 'envi report found'
                envReport = str(item[0])

    if reportNumCount > 2:
        print 'multiple reports detected'
        soilReport = []
        envReport = []
        for item in reportNumbers:
            if str(item[1]) == "SOIL":
                soilReport.append(str(item[0]))
            elif str(item[1]) == "ENVI":
                #print 'envi report found'
                envReport.append(str(item[0]))

    pdfDownload.download_file(outputFolder, soilReport, envReport)


if __name__ == '__main__':
    pdf("CQA1700125")
