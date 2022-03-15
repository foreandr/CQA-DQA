import csv
import SQL_CONNECTOR

def gettingCQAfeeCodes(refno):
    connection = SQL_CONNECTOR.test_connection()
    cursor = connection.cursor()

    query = """SELECT * 
    FROM env_data env
    INNER JOIN report rep
    ON rep.rptno = env.rptno
    WHERE rep.refno = 'CQA2100100' 
    """
    cursor.execute(query)
    feecode = ''
    feecodelist = []
    for item in cursor:
        # print(item[4], item[6])
        feecodelist.append([item[4], item[6]])

    # open and read the file after the appending:
    with open('C:\CQA\FULL CQA - DQA/CQAfeecodes.csv', 'w') as f:
        for i in feecodelist:
            # create the csv writer
            writer = csv.writer(f)

            # write a row to the csv file
            writer.writerow(i)





gettingCQAfeeCodes('CQA2100540')