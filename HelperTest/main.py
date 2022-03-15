import csv
import SQL_CONNECTOR

def gettingCQAfeeCodes(refno):
    connection = SQL_CONNECTOR.test_connection()
    cursor = connection.cursor()
    with open('C:\CQA\FULL CQA - DQA/CQAfeecodes.csv', 'w') as file:
        my_reader = csv.reader(file, delimiter=',')

        #for row in my_reader:
        #    print(row)

        query = """SELECT * 
        FROM env_data env
        INNER JOIN report rep
        ON rep.rptno = env.rptno
        WHERE rep.refno = 'CQA2100540' 
        """



gettingCQAfeeCodes('CQA2100540')