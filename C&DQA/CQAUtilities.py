from openpyxl.styles import Side, Border

import SQL_CONNECTOR
import Colors
import Utilities
import os, sys


def print_dict(my_dict):
    for key, value in my_dict.items():
        print(Colors.bcolors.HEADER + str(key) + ' : ' + str(value) + Colors.bcolors.ENDC)


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
        '21': 'C:N Ratio',
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
        # print(Colors.bcolors.WARNING + 'PRINTING ITEM FROM CURSOR)' + str(item) + Colors.bcolors.ENDC)
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

    # Divide organic carbon by nitrogen
    CNRatioValue = round((Utilities.organicCarbon(available_matter_for_calc) / 0.9) / Nitrogen)
    # print('CNRATIO: ', CNRatioValue)

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
    # print(Colors.bcolors.OKGREEN + 'ENV DICT' + Colors.bcolors.ENDC)
    # print_dict(ENVResult)
    # print(Colors.bcolors.OKGREEN + 'SOIL DICT' + Colors.bcolors.ENDC)
    # print_dict(soilResult)
    # Runs function that merges the two dict's
    tempResult = Utilities.merge_two_dicts(ENVResult, soilResult)

    finalResult = {}
    # print('TEMP RESULT', tempResult)

    cursor.close()
    cnx.close()
    # print '\n', Colors.bcolors.OKCYAN, 'CURRENT VALUES', Colors.bcolors.ENDC  # formatting purposes
    final_array = []
    for key in sorted(tempResult):
        for _key, value in ENVDict.items():
            if key == _key:
                # print _key, value, finalResult[key]
                final_array.append([_key, value, tempResult[key]])  # plS JUST WORK NOW

    # print(final_array)
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
    # print(final_array)  # formatting purposes
    # exit(1)
    index_count = 0
    for i in final_array:
        i.append(index_count)
        index_count += 1
    return final_array, soilResult, ENVResult


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
    print('all items in m3 query ', item_list)

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
        # 2


def DQA_CFIA_FORMATTING(sheet):
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")
    # A. -----------
    for i in range(6, 22):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:L%d' % (i, i), border)

    for i in range(11, 21):  # ARSENIC TO ZINC
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'B%d:L%d' % (i, i), border)

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
    for i in range(25, 29):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:J%d' % (i, i), border)

    border = Border(bottom=thick)  # DOESNT WORK..
    Utilities.FixFormatting(sheet, 'E25:F25', border)

    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'B25:J25', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'B26:J26', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'B27:J27', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B28:J28', border)

    # C. -----------
    for i in range(31, 34):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:J%d' % (i, i), border)

    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'B31:J31', border)
    border = Border(bottom=thick, top=thin)
    Utilities.FixFormatting(sheet, 'B33:J33', border)

    # Minimum Agricultural Values. -----------
    for i in range(50, 54):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:L%d' % (i, i), border)

    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'A50:L50', border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A51:L51', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A52:L52', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A53:L53', border)

    # THE BIG ONE   ----
    for i in range(56, 88):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:L%d' % (i, i), border)

        if (i < 58) or (i > 65 and i < 69):
            continue
        else:
            border = Border(bottom=thin)
            Utilities.FixFormatting(sheet, 'A%d:L%d' % (i, i), border)

    border = Border(bottom=thick, top=thick, right=thick)
    Utilities.FixFormatting(sheet, 'A55:L55', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'D57:D57', border)
    border = Border(left=thin, right=thin)
    Utilities.FixFormatting(sheet, 'E57:E57', border)

    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'A66:L66', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A68:L68', border)

    border = Border(left=thin, right=thin)
    Utilities.FixFormatting(sheet, 'D68:D68', border)

    border = Border(top=thin)
    Utilities.FixFormatting(sheet, 'F68:L68', border)

    border = Border(bottom=thick, right=thick)
    Utilities.FixFormatting(sheet, 'A88:L88', border)


def DQA_ONT_FORMATTING(sheet):
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    # A. --------

    for i in range(6, 21):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:K%d' % (i, i), border)

    for i in range(10, 20):  # ARSENIC TO ZINC
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'B%d:K%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B5:K5', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'I6:K6', border)
    border = Border(right=thin, left=thin)
    Utilities.FixFormatting(sheet, 'D7:D7', border)
    border = Border(right=thin, left=thin)
    Utilities.FixFormatting(sheet, 'D8:D8', border)
    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'G8:K8', border)
    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'F7:F7', border)
    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'F8:F8', border)
    border = Border(right=thin)
    Utilities.FixFormatting(sheet, 'G8:G8', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B9:K9', border)
    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B20:K20', border)

    # B. --------
    for i in range(27, 31):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:K%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'B27:K27', border)

    border = Border(bottom=thin, top=thin)
    Utilities.FixFormatting(sheet, 'B29:K29', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B30:K30', border)

    # C. --------
    for i in range(34, 38):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:K%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'B34:K34', border)

    border = Border(bottom=thin, top=thin)
    Utilities.FixFormatting(sheet, 'B36:K36', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B37:K37', border)

    #  Finished Digestate Quality
    for i in range(52, 55):  # BOX OUTLINE
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:K%d' % (i, i), border)

    border = Border(top=thick)
    Utilities.FixFormatting(sheet, 'A52:K52', border)

    border = Border(top=thick)
    Utilities.FixFormatting(sheet, 'A53:K53', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'D53:D53', border)

    border = Border(bottom=thin, top=thin)
    Utilities.FixFormatting(sheet, 'A54:K54', border)

    border = Border(bottom=thick, right=thick)
    Utilities.FixFormatting(sheet, 'A55:K55', border)

    # FINAL 2 COLUMNS
    for i in range(58, 90):  # BOX OUTLINE
        if i == 68 or i == 67:
            continue
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:K%d' % (i, i), border)

        if i == 58 or i == 59 or i == 69 or i == 70:
            continue
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:K%d' % (i, i), border)

    border = Border(top=thick)
    Utilities.FixFormatting(sheet, 'A58:K58', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A59:K59', border)

    border = Border(left=thin, right=thin)
    Utilities.FixFormatting(sheet, 'D59:D59', border)

    border = Border(right=thin)
    Utilities.FixFormatting(sheet, 'E59:E59', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A67:K67', border)

    border = Border(top=thick)
    Utilities.FixFormatting(sheet, 'A69:K69', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A70:K70', border)

    border = Border(left=thin, right=thin)
    Utilities.FixFormatting(sheet, 'D70:D70', border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'F69:K69', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'D72:D72', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A90:K90', border)

    #fixing the left line
    for i in range(71, 91):
        border = Border(left=thin)
        Utilities.FixFormatting(sheet, 'K%d:K%d'% (i, i), border)

def CQA_ONT_FORMATTING(sheet):
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")

    # A. --------
    for i in range(6, 19):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:H%d' % (i, i), border)

    for i in range(9, 19):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'B%d:H%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'B6:H6', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B8:H8', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B19:H19', border)

    # B. --------
    for i in range(22, 30):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A22:I22', border)

    for i in range(23, 29):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

        border = Border(right=thin)
        Utilities.FixFormatting(sheet, 'G%d:G%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A29:I29', border)

    border = Border(right=thin)
    Utilities.FixFormatting(sheet, 'G29:G29', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'E27:E27', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'E23:E23', border)

    # C. ----
    for i in range(32, 37):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

        border = Border(right=thin, left=thin)
        Utilities.FixFormatting(sheet, 'D%d:D%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A32:I32', border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A34:I34', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A36:I36', border)

    # D. ----
    for i in range(39, 42):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'A39:I39', border)

    border = Border(bottom=thin, top=thin)
    Utilities.FixFormatting(sheet, 'A40:I40', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A41:I41', border)

    # E. -----
    for i in range(45, 48):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'C%d:G%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'C45:G45', border)

    border = Border(bottom=thin, top=thin)
    Utilities.FixFormatting(sheet, 'C46:G46', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'C47:G47', border)

    # Appendix II
    for i in range(52, 62):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'C%d:G%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'C52:G52', border)

    for i in range(53, 61):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'C%d:G%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'C61:G61', border)

    # Reference Compost Quality Parameters for CQA
    for i in range(72, 82):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    for i in range(73, 82):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A72:I72', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A82:I82', border)

    # Appendix III
    for i in range(98, 114):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    for i in range(99, 112):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(top=thin)
    Utilities.FixFormatting(sheet, 'A103:I103', border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A98:I98', border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A113:I113', border)


def CQA_OTHER_FORMATTING(sheet):
    thick = Side(border_style="medium")
    thin = Side(border_style="thin")
    # A --
    for i in range(7, 21):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:H%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'B7:H7', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B9:H9', border)

    for i in range(10, 20):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'B%d:H%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'B20:H20', border)

    # B --
    for i in range(24, 31):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A24:I24', border)

    for i in range(25, 30):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A30:I30', border)

    border = Border(left=thin, right=thin)
    Utilities.FixFormatting(sheet, 'E27:F27', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'D27:D27', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'G26:G26', border)

    border = Border(left=thin)
    Utilities.FixFormatting(sheet, 'G29:G29', border)

    # C --
    for i in range(33, 38):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A33:I33', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A37:I37', border)

    for i in range(34, 37):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    for i in range(34, 38):
        border = Border(right=thin, left=thin)
        Utilities.FixFormatting(sheet, 'D%d:D%d' % (i, i), border)

    # D --
    for i in range(40, 43):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'B%d:I%d' % (i, i), border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A40:I40', border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A41:I41', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'A42:I42', border)

    # ROW 46
    for i in range(46, 64):
        if i == 50 or i == 51 or i == 52 or i == 53:
            pass
        else:
            border = Border(left=thick, right=thick)
            Utilities.FixFormatting(sheet, 'C%d:G%d' % (i, i), border)

    for i in range(47, 63):
        if i == 48 or i == 50 or i == 51 or i == 52 or i == 53 or i == 54:
            pass
        else:
            border = Border(bottom=thin)
            Utilities.FixFormatting(sheet, 'C%d:G%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'C48:G48', border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'C54:G54', border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'C63:G63', border)

    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'C46:G46', border)

    # row 72
    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A72:I72', border)

    for i in range(73, 82):
        border = Border(bottom=thin)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(bottom=thick)
    Utilities.FixFormatting(sheet, 'C82:G82', border)

    # apendix 3
    border = Border(bottom=thick, top=thick)
    Utilities.FixFormatting(sheet, 'A96:I96', border)

    for i in range(97, 110):
        if i == 97 or i == 102 or i == 109:
            pass
        else:
            border = Border(bottom=thin)
            Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A97:I97', border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A102:I102', border)

    border = Border(bottom=thin)
    Utilities.FixFormatting(sheet, 'A109:I109', border)

    border = Border(top=thick, bottom=thick)
    Utilities.FixFormatting(sheet, 'A111:I111', border)

    for i in range(96, 112):
        border = Border(left=thick, right=thick)
        Utilities.FixFormatting(sheet, 'A%d:I%d' % (i, i), border)


def get_Agindex_Phosphorus(CQAREF):
    cnx = SQL_CONNECTOR.test_connection()
    cursor = cnx.cursor()
    query = '''SELECT result_str
    FROM env_data env
    INNER JOIN report rep
    ON rep.rptno = env.rptno
    WHERE env.feecode = 'MPPC380'
    AND refno = '%s'
    ''' % CQAREF
    cursor.execute(query)
    value = 0
    for item in cursor:
        value = float(item[0])
    return value


def get_Agindex_Potassium(CQAREF):
    cnx = SQL_CONNECTOR.test_connection()
    cursor = cnx.cursor()
    query = '''SELECT result_str
    FROM env_data env
    INNER JOIN report rep
    ON rep.rptno = env.rptno
    WHERE env.feecode = 'MKKC380'
    AND refno = '%s'
    ''' % CQAREF
    cursor.execute(query)
    value = 0
    for item in cursor:
        value = float(item[0])
    return value


def get_Agindex_Sodium(CQAREF):
    cnx = SQL_CONNECTOR.test_connection()
    cursor = cnx.cursor()
    query = '''SELECT result_str
    FROM env_data env
    INNER JOIN report rep
    ON rep.rptno = env.rptno
    WHERE env.feecode = 'MNAC380'
    AND refno = '%s'
    ''' % CQAREF
    cursor.execute(query)
    value = 0
    for item in cursor:
        value = float(item[0])
    return value


def agindex_text(number):
    if number < 2:
        return "Salt Injury Probable"
    elif number < 5:
        return "Limit use to soils with excellent drainage and low salt content"
    elif number < 9:
        return "Can be used on soils with poor drainage or high salt content"
    else:
        return "Can be used on all soils"


def similar(a, b):
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()
