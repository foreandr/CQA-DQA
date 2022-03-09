import urllib2, os

def download_file(outputFolder, soil_number = None, env_number = None):
    print "---Downloading PDF---"
    if env_number is not None and isinstance(env_number, str):
        print 'executing third'
        print env_number, 'got this value'
        temp_value = env_number.split("'")
        print temp_value[1]
        env_url = r"https://services.alcanada.com/report-center/envTestReport.rpt?_rptnos=%s,&_tests=A&L-WETDRY,&_hideLogo=false" %(temp_value[1])
        print(env_url)
        print('https://services.alcanada.com/almsrpt/inquiry/printInquiry.do?module=ENVI&rptno=%s' %(temp_value[1]))
        response = urllib2.urlopen(env_url)
        pfile = open(os.path.join(outputFolder,"%s.pdf"%(temp_value[1])), 'wb')
        pfile.write(response.read())
        pfile.close()
        print("%s Download Completed") %(env_number)

    if env_number is not None and isinstance(env_number, list):
        print 'download multi envi'
        print 'executing fourth'
        for item in env_number:
            env_url = r"https://services.alcanada.com/report-center/envTestReport.rpt?_rptnos=%s,&_tests=A&L-WETDRY,&_hideLogo=false" %(item)
            response = urllib2.urlopen(env_url)
            pfile = open(os.path.join(outputFolder,"%s.pdf"%(item)), 'wb')
            pfile.write(response.read())
            pfile.close()
            print("%s Download Completed") %(item)
            
        
if __name__ == "__main__":
    pass
    # download_file("C:\CQABackup", env_number = 'C17362-70051')
