"""NECESSARY IMPORTS"""
from mysql.connector import connect, Error, errorcode
import mysql.connector


def test_connection():
    """WORKING CONFIG"""
    global connection
    config = {
        'user': 'lmsuser',
        'password': 'readonly',
        'host': '10.0.0.26',
        'database': 'alms',
        'buffered': True
    }
    # connects to sql database

    try:
        connection = mysql.connector.connect(**config)
        #print("\nMySQL Database connection successful")
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("Something is wrong with your user name or password")
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            print("Database does not exist")
        else:
            print(err)
    return connection
