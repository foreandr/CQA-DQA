from openpyxl.styles import Side, Border

import SQL_CONNECTOR
import Colors
import Utilities
import os, sys


def OntarioResults(CQARef):
    print Colors.bcolors.OKCYAN + "\nExecuting ON QCResults" + Colors.bcolors.ENDC
    # ------------------------------------------------------Env Report------------------------------------------------------#
    # Dictionary for the environmental report values
    ENVDict = {
        "1": "Arsenic",
        "2": "Cadmium",
        "3": 'Chromium',
        "4": 'Cobalt',
        "5": 'Copper',
        '6': 'Lead',
        "7": 'Mercury',
        '8': 'Molybdenum',
        '9': 'Nickel',
        '10': 'Selenium',
        '11': 'Zinc',
        '12': 'Total FM > 25 mm',
        '13': 'Total sharps > 2.8 mm*',
        '14': 'Total sharps > 12.5 mm',
        '15': 'Respiration-mgCO2-C/g OM/day',
        '16': 'E. coli',
        '17': 'Salmonella spp.',
        '18': 'Total Organic Matter',
        '19': 'Moisture',
        '20': 'Total Organic Matter @ 550 deg C',
        '22': 'C:N Ratio',
        '28': 'Total Solids (as received)',
        '30': 'Bulk Density (As Recieved)',
        '33': 'Ammonia (NH3/NH4-N)',
        '34': 'Total Phosphorus (As P205)',
        '35': 'Total Potassium (as K20)',
        '36': 'Calcium',
        '37': 'Magnesium',
        '38': 'Sulphur',
        '39': 'Total FM > 2.8 mm*',
        '40': 'Total plastics > 2.8 mm*'
    }

    # CoNnEcT To dB
    cnx = SQL_CONNECTOR.test_connection()

    # Querry out the report numbers
    cursor = cnx.cursor()
    query = "SELECT * FROM alms.report WHERE refno='%s'" % (CQARef)
    cursor.execute(query)
    reportNumbers = []
    for item in cursor:
        # print Colors.bcolors.HEADER + "[Ontario Report REF Query:]" + Colors.bcolors.ENDC + str(item)
        tempList = [item[0], item[1]]
        reportNumbers.append(tempList)

    # Store the report numbers
    soilReport = None
    envReport = None
    for item in reportNumbers:
        if str(item[1]) == "SOIL":
            soilReport = str(item[0])
        elif str(item[1]) == "ENVI":
            envReport = str(item[0])

    # print soilReport, envReport

    ENVResult = {}

    for key in ENVDict.keys():
        parameter = ENVDict[key]
        # query every parameter based on name
        if parameter != "Total FM > 25 mm":
            # print key, parameter
            envQuery = r"""select
            ed.result_str
            from env_data as ed
            inner join env_samp as es on ed.rptno=es.rptno and ed.labno=es.labno
            inner join report as r on ed.rptno=r.rptno
            inner join feecode as f on ed.feecode=f.feecode and left(f.rpt_dscrpt,4)<>'TCLP' and f.rpt_flag<>'N'
            left join units as u on ed.unit=u.un_units
            where (ed.result>=0 or ed.result=-3 or ed.result_str='BDL*') and ed.rptno='%s' and f.rpt_dscrpt='%s'
            order by f.prt_sort,f.anly_sort""" % (envReport, parameter)

        # 12
        # query for specific unit of Total FM >25 mm parameter
        elif parameter == "Total FM > 25 mm":
            envQuery = r"""select
                ed.result_str
                from env_data as ed
                inner join env_samp as es on ed.rptno=es.rptno and ed.labno=es.labno
                inner join report as r on ed.rptno=r.rptno
                inner join feecode as f on ed.feecode=f.feecode and left(f.rpt_dscrpt,4)<>'TCLP' and f.rpt_flag<>'N'
                left join units as u on ed.unit=u.un_units
                where (ed.result>=0 or ed.result=-3 or ed.result_str='BDL*') and ed.unit = "pieces/500ml" and ed.rptno='%s' and f.rpt_dscrpt='%s'
                order by f.prt_sort,f.anly_sort""" % (envReport, parameter)

        cursor = cnx.cursor()
        cursor.execute(envQuery)
        # Store the querried information in the ENVResult list
        for item in cursor:
            ENVResult[key] = str(item[0])
        '''
        print Colors.bcolors.BOLD + Colors.bcolors.UNDERLINE + Colors.bcolors.OKBLUE + \
              "Key: " + str(key) + " | " + \
              "Parameter: " + str(parameter) \
              + "" + Colors.bcolors.ENDC
        '''
    # print Colors.bcolors.OKCYAN + 'Ontario----salmonella------' + Colors.bcolors.ENDC
    # print ENVResult['17']
    ##    print '=-----moisture------'
    ##    print ENVResult['19']

    # Querry out moisture to calculate Results as recieved
    envQuery = r"""select
            ed.result
            from env_data as ed
            inner join env_samp as es on ed.rptno=es.rptno and ed.labno=es.labno
            inner join report as r on ed.rptno=r.rptno
            inner join feecode as f on ed.feecode=f.feecode and left(f.rpt_dscrpt,4)<>'TCLP' and f.rpt_flag<>'N'
            left join units as u on ed.unit=u.un_units
            where (ed.result>=0 or ed.result=-3 or ed.result_str='BDL*') and ed.rptno='%s' and f.rpt_dscrpt='Moisture'
            order by f.prt_sort,f.anly_sort""" % (envReport)

    cursor = cnx.cursor()
    cursor.execute(envQuery)
    # Store the querried information
    for item in cursor:
        moisture = item[0]
        ENVResult['19'] = moisture

    for i in range(34, 39):
        temp = str(i)
        dm = 100 - float(moisture)
        dm = dm / 100
        ENVResult[temp] = str(float(ENVResult[temp]) * dm)
        # ------------------------------------------------------Sieves------------------------------------------------------#
    # 22
    # List of all possible Sieve sizes
    sieveList = [r'Sieve 2 Inch (% Passing)', r'Sieve 1 Inch (% Passing)', r'Sieve 1/2 Inch (% Passing)',
                 r'Sieve 3/8 Inch (% Passing)',
                 r'Sieve 1/4 Inch (% Passing)']

    sieveDict = {}
    sieveResult = {}
    targetPercent = 80.0
    smallestParameter = None
    # Querry out the sieve sizes
    for item in sieveList:
        parameter = item
        envQuery = r"""select
        ed.result_str
        from env_data as ed
        inner join env_samp as es on ed.rptno=es.rptno and ed.labno=es.labno
        inner join report as r on ed.rptno=r.rptno
        inner join feecode as f on ed.feecode=f.feecode and left(f.rpt_dscrpt,4)<>'TCLP' and f.rpt_flag<>'N'
        left join units as u on ed.unit=u.un_units
        where (ed.result>=0 or ed.result=-3 or ed.result_str='BDL*') and ed.rptno='%s' and f.rpt_dscrpt='%s'
        order by f.prt_sort,f.anly_sort""" % (envReport, parameter)
        cursor = cnx.cursor()
        cursor.execute(envQuery)
        # Goes through all sieve sizes
        for item in cursor:
            a = str(item[0])
            item = float(item[0])
            # Check to see if the current sieve size is below the target size and if it is then sets the value too 999
            if item - targetPercent <= 0:
                sieveDict[parameter] = 999
                sieveResult[parameter] = a
                # If its not below the target size then sets the value to the difference of the two numbers
            else:
                sieveDict[parameter] = item - targetPercent
                sieveResult[parameter] = a

    # get smallest size
    for key, value in sorted(sieveDict.iteritems(), key=lambda (k, v): (v, k)):
        smallestParameter = key
        break

    tempSieve = smallestParameter.replace('Sieve', '').replace('(% Passing)', '')[1:-1]

    # sets the sieve size
    ENVResult['22'] = tempSieve

    # ------------------------------------------------------Soil Report------------------------------------------------------#

    soilResult = {}

    # Querrys out the different soil report values
    # 20 and 29: PH
    soilResult['20'] = Utilities.getPH(CQARef)
    soilResult['29'] = Utilities.getPH(CQARef)

    # 23: salt
    query = """select
        IF(s.salt>0,s.salt,null) as salt
        from soil as s
        inner join report as r on r.rptno=s.rptno
        where s.rptno='%s'
        order by s.labno""" % (soilReport)

    cursor = cnx.cursor()
    cursor.execute(query)
    for item in cursor:
        soilResult['23'] = str(round(item[0], 1))

    soilResult['32'] = Utilities.getNitrogen(CQARef)

    # -----------Total Organic matter-------------#
    # print Colors.bcolors.OKCYAN + "----------------------------------Ontario total organic-------------------------------" + Colors.bcolors.ENDC

    soilResult['32'] = Utilities.getNitrogen(CQARef)  # Nitrogen
    soilResult['18'] = Utilities.getTotalOrganicMatter(CQARef)  # Total organic MAtter
    available_matter_for_calc = Utilities.getAvailableOrganicMatter(CQARef)
    totalOrganicCarbon2 = Utilities.organicCarbon(available_matter_for_calc)
    Nitrogen = float(soilResult['32'])
    # print 'Total Organic            :' + str(soilResult['18'])
    # print 'AVailable Organic        : ' + str(available_matter_for_calc)
    # print 'OG CARB * 0.6            : ' + str(totalOrganicCarbon2)
    # print 'Nitrogen                 :' + str(Nitrogen)

    # Divide organic carbon by nitrogen
    CNRatioValue = round((Utilities.organicCarbon(available_matter_for_calc) / 0.9) / Nitrogen)
    print('CNRATIO: ', CNRatioValue)

    # print 'CNRatioValue             :' + str(CNRatioValue)

    cNRatio = str("%d:1" % (CNRatioValue))
    # print 'Calculated CN Ratio = ' + cNRatio
    soilResult['21'] = cNRatio
    soilResult['31'] = cNRatio

    # ----------------------------Interpretation--------------------------------------------#

    # CEC Calculations (k_m3, mg_m3, ca_m3, na, buffer)
    # 24-27
    CECDict = {
        'k_m3': [],
        'mg_m3': [],
        'ca_m3': [],
        'na': [],
        'buffer': []
    }
    # Querry the cecDict
    for key in CECDict.keys():
        parameter = key
        query = """select
            IF(s.%s>0,s.%s,null) as %s
            from soil as s
            inner join report as r on r.rptno=s.rptno
            where s.rptno='%s'
            order by s.labno""" % (parameter, parameter, parameter, soilReport)

        cursor = cnx.cursor()
        cursor.execute(query)
        # If they don't exist then stop else add them to the dict
        for value in cursor:
            if value is None:
                sys.exit()
            else:
                CECDict[key].append(float(value[0]))

    # perk - k_m3, 390, E28
    # perMG - _mg_m3, 121.6, E29
    # perCa - ca_m3, 200.0, E30
    # perNa - na, 230.0., E26

    # Call the cakculateCEC function
    cec = Utilities.calculateCEC(CECDict)

    # calculation parameters
    interpDict = {
        '24': ['na', 230.0],
        '25': ['k_m3', 390.0],
        '26': ['mg_m3', 121.6],
        '27': ['ca_m3', 200.0]
    }

    # add results to soilResults
    for key in interpDict.keys():
        # call def percentCalc(cec, paramterName, y, CECDict):
        result = Utilities.percentCalc(cec, CECDict[interpDict[key][0]][0], interpDict[key][1], CECDict)
        soilResult[key] = result

        # ----------------------------Merging and Formatting--------------------------------------------#

    # Runs function that merges the two dict's
    tempResult = Utilities.merge_two_dicts(ENVResult, soilResult)

    finalResult = {}
    # Goes through all the results
    for key in tempResult.keys():
        # Stores results into lists
        value = tempResult[key]
        digits = Utilities.formatDict[key]

        # Sees what format the number needs and modifies the value to fit the format
        try:
            float(value)
            if digits == 2:
                finalValue = "%.2f" % (float(value))
                # print key, value, finalValue
                finalResult[key] = finalValue
            elif digits == 1:
                finalValue = "%.1f" % (float(value))
                # print key, value, finalValue
                finalResult[key] = finalValue
            elif digits == 0:
                finalValue = int(value)
                finalResult[key] = finalValue
            elif digits == '2%':
                finalValue = "%.2f" % (float(value)) + '%'
                finalResult[key] = finalValue

        # If there's a value error then don't change it
        except ValueError:
            finalResult[key] = value

    cursor.close()
    cnx.close()
    # print 'ON QC Final Result ' + str(finalResult)

    final_array = []
    for key in sorted(finalResult):
        for _key, value in ENVDict.items():
            if key == _key:
                # print _key, value, finalResult[key]
                final_array.append([_key, value, finalResult[key]])  # plS JUST WORK NOW

    def vector2d_sort(array):
        for i in range(len(array) - 1):
            # print i, array[i]
            if int(array[i][0]) > int(array[i + 1][0]):
                biggerTemp = array[i]
                smallerTemp = array[i + 1]
                array[i] = smallerTemp
                array[i + 1] = biggerTemp

                vector2d_sort(array)

    vector2d_sort(final_array)
    count = 0
    for i in final_array:
        i.append(count)
        count += 1
    final_array.append(['29', 'REAL C:N RATIO', int(CNRatioValue), count])
    print '\n', Colors.bcolors.OKCYAN, 'CURRENT VALUES', Colors.bcolors.ENDC  # formatting purposes
    for i in final_array:
        print(i)
    print('\n')  # formatting purposes
    return final_array


def getOtherResults(CQAREF):
    print(CQAREF)

    cnx = SQL_CONNECTOR.test_connection()
    cursor = cnx.cursor()
    query = '''
    select salt, perna_m3, perk_m3, permg_m3, perca_m3
    from agdata a
    inner join report r
    on r.rptno = a.rptno
    inner join soil s
    on s.rptno = a.rptno
    where route_4 = '%s'
    ''' % CQAREF
    cursor.execute(query)
    item_list = []
    for item in cursor:
        item_list.append(item)
    print('all items in m3 query ' , item_list)

    new_list = map(float, item_list[0])

    newDict = {}
    # perna_m3, perk_m3, permg_m3, perca_m3
    newDict['salt'] = new_list[0]
    newDict['perna_m3'] = new_list[1]
    newDict['perk_m3'] = new_list[2]
    newDict['permg_m3'] = new_list[3]
    newDict['perca_m3'] = new_list[4]
    print(newDict)
    return newDict


def get_dry_matter(CQAREF):
    cnx = SQL_CONNECTOR.test_connection()
    cursor = cnx.cursor()
    query = '''
    SELECT result_str
    FROM env_data env
    INNER JOIN report rep
    ON rep.rptno = env.rptno
    WHERE env.feecode = 'GTSZ280'
    AND refno = '%s'
    ''' % CQAREF
    cursor.execute(query)
    value = 0
    for item in cursor:
        value = float(item[0])
    return value


def get_partcile(CQAREF):
    cnx = SQL_CONNECTOR.test_connection()
    cursor = cnx.cursor()
    query = '''
    SELECT feecode, result_str
    FROM env_data env
    INNER JOIN report rep
    ON rep.rptno = env.rptno
    WHERE rep.refno = '%s' AND env.feecode = 'SQCC023'
    OR rep.refno = 'CQA2200061' AND env.feecode = 'SQBC023'  
    OR rep.refno = 'CQA2200061' AND env.feecode = 'SSCC023' 
    OR rep.refno = 'CQA2200061' AND env.feecode = 'SSBC023'
    OR rep.refno = 'CQA2200061' AND env.feecode = 'SQAC023'  
    ''' % CQAREF
    cursor.execute(query)
    usingDict = {}
    for item in cursor:
        usingDict[item[0]] = item[1]

    print('\n')
    for key, value in usingDict.items():
        print(key, value)

    if usingDict['SQCC023'] >= 79.5:
        return '1/4'
        # 1/4
    elif usingDict['SQBC023'] >= 79.5:
        return '3/8'
        # 3/8
    elif usingDict['SSCC023'] >= 79.5:
        return '1/2'
        # 1/2
    elif usingDict['SSBC023'] >= 79.5:
        return '1'
        # 1
    else:
        return '2'
        #2


def DQA_CFIA_FORMATTING(sheet):
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")
    # A. -----------
    for i in range(6, 22): # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:L%d' % (i, i), border)

    for i in range(11, 21): # ARSENIC TO ZINC
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'B%d:L%d' %(i, i) , border)

    for i in range(7, 10):
        border = Border(top=thin, left=thin)
        Utilities.FixFormatting(sheet, 'D%d:L%d' % (i, i), border)
        Utilities.FixFormatting(sheet, 'E%d:L%d' % (i, i), border)
        Utilities.FixFormatting(sheet, 'G%d:L%d' % (i, i), border)
        Utilities.FixFormatting(sheet, 'I%d:L%d' % (i, i), border)

    border = Border(top=thick)
    Utilities.FixFormatting(sheet, 'B6:L6', border)
    border = Border(top=thin)
    Utilities.FixFormatting(sheet, 'D10:L10', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B10:L10', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B21:L21', border)

    # B. -----------
    for i in range(25, 29): # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:J%d' % (i, i), border)


    border = Border(bottom=thick) #DOESNT WORK..
    Utilities.FixFormatting(sheet, 'E25:F25', border)
    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'B25:J25', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'B26:J26', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'B27:J27', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B28:J28', border)




#Utilities.FixFormatting(sheet, 'B10:L10' , border)