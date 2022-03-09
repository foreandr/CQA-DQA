import urllib2, os

def download_file(outputFolder, soil_number = None, env_number = None):
    print "---Downloading PDF---"
    if soil_number is not None and isinstance(soil_number, str):
        soil_url = r"https://services.alcanada.com/report-center/soilTestReport.rpt?_rptnos=%s,&_hideSigns=true,&_tests=SQA_COMPOST,&_miscs=true,&_hideLogo=false" %(soil_number)
        response = urllib2.urlopen(soil_url)
        pfile = open(os.path.join(outputFolder, "%s.pdf"%(soil_number)), 'wb')
        pfile.write(response.read())
        pfile.close()
        print("%s Download Completed") %(soil_number)

    if soil_number is not None and isinstance(soil_number, list):
        for item in soil_number:
            soil_url = r"https://services.alcanada.com/report-center/soilTestReport.rpt?_rptnos=%s,&_hideSigns=true,&_tests=SQA_COMPOST,&_miscs=true,&_hideLogo=false" %(item)
            response = urllib2.urlopen(soil_url)
            pfile = open(os.path.join(outputFolder, "%s.pdf"%(item)), 'wb')
            pfile.write(response.read())
            pfile.close()
            print("%s Download Completed") %(item)


    

    if env_number is not None and isinstance(env_number, str):
        env_url = r"https://services.alcanada.com/report-center/envTestReport.rpt?_rptnos=%s,&_tests=AL-CQA2,&_hideLogo=false" %(env_number)
        response = urllib2.urlopen(env_url)
        pfile = open(os.path.join(outputFolder,"%s.pdf"%(env_number)), 'wb')
        pfile.write(response.read())
        pfile.close()
        print("%s Download Completed") %(env_number)

    if env_number is not None and isinstance(env_number, list):
        print 'download multi envi'
        for item in env_number:
            env_url = r"https://services.alcanada.com/report-center/envTestReport.rpt?_rptnos=%s,&_tests=AL-CQA2,&_hideLogo=false" %(item)
            response = urllib2.urlopen(env_url)
            pfile = open(os.path.join(outputFolder,"%s.pdf"%(item)), 'wb')
            pfile.write(response.read())
            pfile.close()
            print("%s Download Completed") %(item)
            
        
if __name__ == "__main__":
    download_file("C:\CQABackup", env_number = 'C17362-70051')
