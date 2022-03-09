# -*- coding: utf-8 -*-

import Tkinter
import tkSimpleDialog
import tkMessageBox
import mysql.connector
import openpyxl
from time import ctime, sleep
import urllib2
from os import mkdir, path
import os
from exceptions import *



class MySQLConnector:
    """Simple wrapper of mysql-connector functionality"""
    def __init__(self, config):
        """
        Creates a new wrapper connector to the mysql server specified by the ``config``

        Parameters
        ----------
        config : `dict`
            A dictionary of the config to log in with
        ``"host"`` : `str`
            The IP of the server to connect to (ei. 192.168.0.13)
        ``"database"`` : `str`
            The name of the database to connect to
        ``"user"`` : `str`
            The name of the user to log on as (ei. root)
        ``"password"`` : `str`
            The password to use went attempting to connect (ei. password1)
        ``"buffered"`` : `bool`
            TODO: I forget

        Notes
        -----
        self.cnx is None on failure, case check it
        """
        try:
            self.cnx = mysql.connector.connect(**config)
        except mysql.connector.Error as E:
            self.err = str(E)
            self.cnx = None
        return
    
    # Creates a mysql cursor and executes the query
    # Doesn't do any query checking/parsing
    def Query(self, query, dictionary=False):
        """
        Makes a query with the given ``query`` and returns the result

        Parameters
        ----------
        query : `str`
            The query to send to MySql
        dictionary : `bool`
            Defaults to false.
            If false, returns a list of lists as per usual.
            If true, retuns a list of dictionaries, where the key is the field

        Returns
        -------
        result : `list` [`list`]
            The result of the query, a list of rows.
        """
        cursor = self.cnx.cursor(buffered=True, dictionary=dictionary)
        cursor.execute(query)
        return list(cursor)
    
    # Closes the connection, y'better call this or Imma come for you
    def Close(self):
        """Closes the connection to the MySql server"""
        self.cnx.close()
        del self

MYSQL_CONFIG = {
    "user":     "lmsuser",
    "password": "readonly",
    "host":     "10.0.0.26",
    "database": "alms",
    "buffered": True
}

def QueryRefnoExists(sql, referal_code):
    """Queries the refno to check that it actually exists"""
    query = """
        SELECT refno FROM alms.report
        WHERE refno = "%s";
    """ % referal_code
    result = sql.Query(query)
    if len(result)==0:
        return False
    return True

def QueryReportNumbers(sql, referal_code):
    """
        Returns the Enviromental Report Code AND the Soil Report Code of the report with the same given ``referral_code``.
        
        Parameters
        ----------
        sql : `MySQLConnector`
            The connector to the database, should be already connected.
        referal_code : `str`
            The 'refno' used to refer to the report in question.
        
        Returns
        -------
        envRpt, soilRpt : `tuple` [`str`, `str`]
            The envRpt is the code for the enviromental report
            The soilRpt is the code for the soil report
            Prints a messege and returns (None, None) if the query fails
    """
    query = "SELECT rptno FROM alms.report WHERE refno='%s' ORDER BY module LIMIT 0, 2;" % referal_code
    reports = sql.Query(query)
    if len(reports)!=2:
        return (None, None)
    envi_rpt_code = reports[0][0]
    soil_rpt_code = reports[1][0]
    return (envi_rpt_code, soil_rpt_code)

def QueryFullAddress(sql, referal_code):
    """
        Grabs the full address of the company the report is for, using the given ``referal code``.

        Parameters
        ----------
        sql : `MySQLConnector`
            The connector to the database, should be already connected.
        referal_code : `str`
            The 'refno' used to refer to the report in question.

        Return
        ------
        full_address : `tuple` [`str`, `str`]
            The address in two parts, in a double tuple,
            The first part is the civic: number, street, and PO box.
            The second part is the region: city, province/state, and postal/zip code.
            Prints a messege and returns ("[Not Found]", "") if the query fails
    """
    query = """
        SELECT address1, address2, city, state, zip FROM alms.report
        WHERE refno = "%s";
    """ % referal_code
    result = sql.Query(query, dictionary=True)
    if len(result)==0:
        return ("[Not Found]", "")
    result = result[0]
    civic_address = result['address1'] +" "+ result['address2']
    region_address = result['city'] +" "+ result['state'] +" "+ result['zip']
    return (civic_address, region_address)

def QueryCompanyInfo(sql, referal_code):
    """
        Grabs the company, attention, and name, from the report with the given ``referal_code``.

        Parameters
        ----------
        sql : `MySQLConnector`
            The connector to the database, should be already connected.
        referal_code : `str`
            The 'refno' used to refer to the report in question.
            
        Return
        ------
        info : `tuple` [`str`, `str`, `str`]
            The return is a tuple of the 3 pieces of info.
            (company, attention, name)
            Prints a messege and returns ("Not Found" * 3) if the query fails
    """
    query = """
        SELECT company, attn, grow_1 FROM alms.report
        WHERE refno = "%s" LIMIT 0, 1;
    """ % referal_code
    result = sql.Query(query)
    # Case check query was successful
    if len(result)==0:
        return ("[Not Found]", "[Not Found]", "[Not Found]")
    # Return the (company, attention, name)
    return (result[0][0], result[0][1], result[0][2])

def GrabTextureData(sql, soil_report_code):
    """
        Grabs the soil texture data and returns it as a dictionary.
        { 'stclass', stsand', 'stsilt', 'stclay' }

        Parameters
        ----------
        sql : `MySQLConnector`
            The connector to the database, should be already connected.
        soil_report_code : `str`
            The 'rptno' of the soil report. Can be gotten using QueryReportNumbers()

        Return
        ------
        texture_data : `dict`
            A dictionary of the texture data.
            Dictionary is empty if something goes wrong.
            Keys: 'stclass', 'stsand', 'stsilt', 'stclay'
    """
    
    query = """
        SELECT feecode, result_str FROM alms.agdata
        WHERE rptno = "%s";
    """ % soil_report_code
    results = sql.Query(query)
    texture_data = {}

    # Go through each record and add the result to the dictionary with the module as the key
    for result in results:
        texture_data[result[0].lower()] = result[1]

    if len(texture_data)==0:
        print("[WARNING]: Failed to query the soil texture data!")
        print(" . . . . : Please manually review the class, sand, silt, and clay info boxes")

    return texture_data

def GrabSoilData(sql, soil_report_code):
    """
        Queries and grabs all relevent soil data with the given ``soil_report_code``

        Parameters
        ----------
        sql : `MySQLConnector`
            The connector to the database, should be already connected.
        soil_report_code : `str`
            The 'rptno' of the soil report. Can be gotten using QueryReportNumbers()

        Return
        ------
        soil_data : `dict`
            A dictionary of all the soil data.
            The keys are the same spelling as the fields in the sql table.
    """            
    query = """
        SELECT id_1, rptno, om, p1, perp, k, perk, mg, permg, ca, perca, na, perna, ph
        FROM soil WHERE soil.rptno="%s";
    """ % soil_report_code
    result = sql.Query(query, dictionary=True)
    if len(result)==0:
        return {}
    for key in result[0]:
        result[0][key] = str(result[0][key])
    return result[0]

def GetReportDate():
    """
        Returns the current date, formatted, to be used on the report

        Return
        ------
        data : `str`
            The formatted date. e.i."Apr 29, 2019"
    """
    full_date = ctime()
    # Example full date:    "Mon Apr 29 09:57:27 2019"
    day = full_date[4:10] # "Apr 29"
    year = full_date[20:] # "2019"
    return day + ", " + year   # "Apr 29, 2019"

def download_soil_pdf(output, soil_number):
    """Returns an exception object on failure, None on success"""
    soil_url = r"https://services.alcanada.com/report-center/soilTestReport.rpt?_rptnos=%s,&_hideSigns=false,&_tests=_,&_miscs=true" % (soil_number)
    try:
        response = urllib2.urlopen(soil_url)
    except ConnectionError as E:
        return E
    file_name = path.join(output, soil_number + ".pdf")
    pfile = open(file_name, 'wb')
    data = response.read()
    if len(data)==0:
        return "Downloaded an empty file"
    try:
        pfile.write(data)
    except PermissionError as E:
        pfile.close()
        return E
    pfile.close()
    return None

def download_env_pdf(output, env_number):
    """Returns an exception object on failure, None on success"""
    env_url = r"https://services.alcanada.com/report-center/envTestReport.rpt?_rptnos=%s,&_tests=" %(env_number)
    try:
        response = urllib2.urlopen(env_url)
    except ConnectionError as E:
        return E
    file_name = path.join(output, env_number + ".pdf")
    pfile = open(file_name, 'wb')
    data = response.read()
    if len(data)==0:
        return "Downloaded an empty file"
    try:
        pfile.write(data)
    except PermissionError as E:
        pfile.close()
        return E
    pfile.close()
    return None

def main():
    # Simple counter for tracking whether an error has occured,
    # if so, we ask the user if they want to retry
    errors_occured = 0

    # The root window of the GUI which we need although don't use
    # did I mention how much I dislike tkinter?
    root = Tkinter.Tk()
    
    # Centre the window
    root.geometry("+800+400")

    # This will be the window title for all the popups
    title = "STP Report Automizer"

    # Hide the root window because we're not using it
    root.withdraw()
    
    # This will be the connector for the sql database
    sql = MySQLConnector(MYSQL_CONFIG)

    # Case check success
    if sql.cnx == None:
        tkMessageBox.showerror(title, "Couldn't connect to the MySql database!\n\n"+sql.err)
        root.destroy() # we made a window, gotta delete it
        return -1

    # The name of the template workbook
    workbook_name = "Master.xlsx"

    # Load the workbook we need to edit
    try:
        Workbook = openpyxl.load_workbook(workbook_name)
        Worksheet = Workbook.active # Grabs first sheet or active sheet
    except BaseException as E:
        tkMessageBox.showerror(title, "Failed to open the template excel file 'Master'\n\n" + str(E))
        root.destroy() # we made a window, gotta delete it
        sql.Close() # we made a connection, gotta close it
        return -1

    # Loop as long as the user wants to do more reports
    # Break the loop when the user is done
    # Continue the loop if an error occurs and you need try again
    # From now on, break instead of returning, as breaking will still call cleanup
    while True:
        
        # Ask the user for the referal code
        referal_code = tkSimpleDialog.askstring(title, "Please enter in the referal number") #this is how enter key works

        # Case check Cancel or Quit
        if referal_code==None: break
        
        # Bring it to all uppercase
        referal_code = referal_code.upper()

        # Case check empty input, try again
        if referal_code=="": continue
        
        # Check whether the requested report exists
        if not QueryRefnoExists(sql, referal_code):
            if tkMessageBox.askretrycancel(title, "That referal code doesn't exist.\nWould you like to try again?"):
                continue # user tries again
            break # user cancels and quits

        # # # # # # # # # # # # # # # # # # # # #
        #   QUERY ALL THE INFORMATION WE NEED   #
        # # # # # # # # # # # # # # # # # # # # #

        # Query the enviro and soil report codes from the referal code
        env_code, soil_code = QueryReportNumbers(sql, referal_code)

        # Case check success, error otherwise
        if env_code is None:
            if tkMessageBox.askretrycancel(title, "[ERROR] Failed to query the enviro and soil report numbers!\nWould you like to try again?"):
                continue # user retries
            break # user cancels and quits
        
        # Query the company, attn, and name from the referal code
        company, attention, name = QueryCompanyInfo(sql, referal_code)
        
        # Warn the user if it fails
        if company == "[Not Found]":
            tkMessageBox.showwarning(title, "Failed to query the company, attn, and name.")
            errors_occured += 1

        # Query the full address of the company from the referal code
        civic_address, region_address = QueryFullAddress(sql, referal_code)

        # Warn the user if it fails
        if civic_address == "[Not Found]":
            tkMessageBox.showwarning(title, "Failed to query the full address.")
            errors_occured += 1

        # Query the texture data from the soil report
        texture_data = GrabTextureData(sql, soil_code)
        # Your keys are 'class', 'clay', 'silt', 'sand'

        # Warn the user if it fails
        if len(texture_data)==0:
            tkMessageBox.showwarning(title, "Failed to query the soil texture data.")
            errors_occured += 1

        # Query the soil data from the soil report
        soil_data = GrabSoilData(sql, soil_code)
        # Dictionary keys are the same spelling as the sql field

        # Warn the user if it fails
        if len(soil_data)==0:
            tkMessageBox.showwarning(title, "Failed to query the soil data.")
            errors_occured += 1

        # # # # # # # # # # # # # # # # # # # # # # #
        #   PLACE ALL THE INFORMATION IN THE EXCEL  #
        # # # # # # # # # # # # # # # # # # # # # # #

        # The default value to give cells if there's none
        default = "[Not Found]"

        # Header
        Worksheet['B4'] = GetReportDate() # Set the Report Date
        Worksheet['B6'] = company # Set the company
        Worksheet['B7'] = civic_address # Set the street address
        Worksheet['B8'] = region_address # Set the region address
        Worksheet['B10'] = attention # Set the attention
        Worksheet['B11'] = name # Set the name of whoever
        Worksheet['B12'] = soil_data.get('id_1', default) # Set the sample id at header

        # Texture Classification
        Worksheet['E14'] = soil_code # Set the soil report number
        Worksheet['B17'] = soil_data.get('id_1', default) # Set the sample id
        Worksheet['E17'] = texture_data.get('stsand', default) # Set the sand%
        Worksheet['G17'] = texture_data.get('stsilt', default) # Set the silt%
        Worksheet['I17'] = texture_data.get('stclay', default) # Set the clay%
        
        # Herbicides
        Worksheet['E19'] = env_code # Set the environment report code
        
        # Fertility Analysis Overview
        Worksheet['E26'] = soil_code # Set the soil report code
        
        # Organic Matter
        Worksheet['C42'] = soil_data.get('om', default) # Set the organic matter
        
        # Phosphorus
        Worksheet['C47'] = soil_data.get('p1', default) # Set the phosphorus
        Worksheet['F47'] = soil_data.get('perp', default) # Set the phosphorus percent
        
        # Potassium
        Worksheet['C53'] = soil_data.get('k', default) # Set the potassium
        Worksheet['F53'] = soil_data.get('perk', default) # Set the potassium percent
        
        # Magnesium
        Worksheet['C58'] = soil_data.get('mg', default) # Set the magnesium
        Worksheet['F58'] = soil_data.get('permg', default) # Set the magnesium percent
        
        # K/Mg Ratio
        k = soil_data.get('perk', None) # Get the potassium percent
        mg = soil_data.get('permg', None) # Get the magnesium percent
        try:
            kmg_ratio = str(round(k / mg, 2)) # Divide to 2 decimal places
        except TypeError:
            kmg_ratio = default
        # Set the ratio
        Worksheet['C63'] = kmg_ratio
        
        # Calcium
        Worksheet['C68'] = soil_data.get('ca', default) # Set the calcium
        Worksheet['F68'] = soil_data.get('perca', default) # Set the calcium percent
        
        # Sodium
        Worksheet['C72'] = soil_data.get('na', default) # Set the sodium
        Worksheet['F72'] = soil_data.get('perna', default) # Set the sodium percent
        
        # pH
        Worksheet['C77'] = soil_data.get('ph', default) # Set the ph

        # Reccomendations
        Worksheet['B104'] = soil_data.get('id_1', default) # Set the id
        Worksheet['B110'] = soil_data.get('id_1', default) # Set the id

        # # # # # # # # # # # # # #
        #    SAVE THE NEW EXCEL   #
        # # # # # # # # # # # # # #

        # The path where all related output files go to
        output_path = path.join("Reports", referal_code)
        
        # Make the folder if it doesnt exist
        try:
            mkdir(output_path)
        except WindowsError as E:
            tkMessageBox.showwarning(title, "The folder for this report already exists.\nThe report might've already been ran through the program.\n\n"+str(E))

        # The save path of our workbook is the reports folder, its own folder, then the report number
        save_name = path.join(output_path, "Report_"+referal_code+".xlsx")
        
        # Attempt to save the workbook
        while True:
            try:
                Workbook.save(save_name) # Save our changes
                break
            except PermissionError as E:
                if tkMessageBox.askretrycancel(title, str(E)+"\n\nYou probably have the excel open in another window, please close it and try again."):
                    continue
                errors_occured += 1
                break

        # # # # # # # # # # # # # #
        #    DOWNLOAD THE PDFs    #
        # # # # # # # # # # # # # #

        # Attempt to download the soil report pdf
        state = download_soil_pdf(output=output_path, soil_number=soil_code)
        
        # Warn the user if the download failed
        if state != None:
            tkMessageBox.askretrycancel(title, str(state)+"\n\nIf it's a connection thing, check if you're offline, if it's a permission thing, check if you have the would be pdf open in another window and close it.")
            errors_occured += 1

        # Attempt to download the enviro report pdf
        state = download_env_pdf(output=output_path, env_number=env_code)

        # Warn the user if the download failed
        if state != None:
            tkMessageBox.askretrycancel(title, str(state)+"\n\nIf it's a connection thing, check if you're offline, if it's a permission thing, check if you have the would be pdf open in another window and close it.")
            errors_occured += 1

        # # # # # # # # # # # # #

        # Warn the user that there were errors
        if errors_occured > 0:
            tkMessageBox.showinfo(title, "%i many errors had occured, you can either retry or manually review the missed values in the report." % errors_occured)

        # Ask the user if they want to do another report
        if tkMessageBox.askyesno(title, "Finished! Would you like to do another one?"):
            continue # user does another document
        break # user quits the program

    # End of loop, cleanup!
    sql.Close()
    Workbook.close()
    root.destroy()
    return 0

if __name__ == "__main__": main()
