import SQL_CONNECTOR
import csv
def gettingfeeCodes(refno):
    '''REPLACE LABBNO WITH INPUT, REPLACE REFNO WITH ENTERED REFNO'''
    listOfCodes = []
    listOfNames = []
    full_usable_list = []
    with open('C:\CQA\FULL CQA - DQA/feecodescv.csv') as file:
        my_reader = csv.reader(file, delimiter=',')
        for row in my_reader:
            #print row[0], row[1]
            listOfCodes.append(row[0])
            listOfNames.append(row[1])

    for i in range(0, len(listOfCodes)):
        #print listOfCodes[i]
        cnx = SQL_CONNECTOR.test_connection()
        cursor = cnx.cursor()
        query = """SELECT * 
        FROM env_data env
        INNER JOIN report rep
        ON rep.rptno = env.rptno
        WHERE rep.refno = '%s' and feecode = '%s'
        """ % ((refno), (listOfCodes[i]))
        cursor.execute(query)
        feecode = ''
        for item in cursor:
            #print listOfCodes[i], item
            feecode = str(item[6])
            #print feecode
        templist = [i, listOfNames[i], listOfCodes[i], feecode]
        full_usable_list.append(templist)
    return  full_usable_list
#
#feecodes_and_values = gettingfeeCodes()
#for i in feecodes_and_values:
#    print i
