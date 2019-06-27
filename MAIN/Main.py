# Import List
from __future__ import print_function
from datetime import datetime
from os import listdir
from os.path import isfile, join
import smtplib
import glob
import logging
import logging.config
import os, os.path
import shutil
import os.path
import win32com.client

#Directory Variables
tobeloaded = r"\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\_to be loaded/"
datasheetloading = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\Datasheet Loading/'
temp = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\temp'
completed = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\Completed Load Files/'
logdir = r'M:\GlobalImageManagement\Datasheet Loading New\Log'

#Glob Dirs
subtobeloaded = r'//wfsrvgbco001003/Datasrv5/MPP/GlobalImageManagement/Datasheet Loading New/_to be loaded/*'
subdatasheetloading = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\Datasheet Loading/*'


#Create Log File
logging.basicConfig(format='%(levelname)s: %(asctime)s: %(message)s', filename="Log", level=logging.INFO)
logging.info("Log File Created")

#Pull Files from 'to_be_loaded' - oldest files first
list_of_files = glob.glob(subtobeloaded)
oldest_file = max(list_of_files, key=os.path.getctime)
print (oldest_file)
shutil.move(oldest_file, datasheetloading)
logging.info("Oldest File in to_be_loaded folder has been moved to Datasheet Loading")

# Create Loop
## used to allow the program to retry the process if there were over 2000 articles on the previous run
restart = 1
any_files = 0
while restart < 10:
    #Launch Excel and Execute Macros
    try:
        wb.Close(False)
        xlApp.Quit()
        del xlApp
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser('M:\GlobalImageManagement\Datasheet Loading New\Merge Spreadsheet.xlsm')
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('simpleXlsMerger')
        logging.info("Merger Ran")
        xlApp.Run('Clean_Sort')
        logging.info("Clean_Sort Ran")
        xlApp.Run('datavalidation')
        logging.info("Article Number Check Ran")
    except:
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser('M:\GlobalImageManagement\Datasheet Loading New\Merge Spreadsheet.xlsm')
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('simpleXlsMerger')
        logging.info("Merger Ran")
        xlApp.Run('Clean_Sort')
        logging.info("Clean_Sort Ran")
        xlApp.Run('datavalidation')
        logging.info("Article Number Check Ran")




    
    #Check for <2000
    f=open("number.txt","r", encoding="utf-16")
    number=(f.read())

    #Stop Loop if under Article Limit
    if int(number) <1001:
        if int(number) >999:
            logging.info("Articles Found:")
            logging.info(number)
            restart = 11

        else:
            if len(os.listdir(tobeloaded) ) == 0:
                logging.info("All Load files Loaded")
                restart = 11
        
            else:
                list_of_files = glob.glob(subtobeloaded)
                oldest_file = max(list_of_files, key=os.path.getctime)
                print (oldest_file)
                shutil.move(oldest_file, datasheetloading)
                logging.info("Gathered another Load File")
            



    else:
        list_of_files = glob.glob(subdatasheetloading)
        latest_file = min(list_of_files, key=os.path.getctime)
        print (latest_file)
        shutil.move(latest_file, temp)
        logging.info("Over 1000 Article Limit, removed offending load file")

        restart = restart + 1
        if restart == 10:
            logging.critical("Program has re-run 10 times, and there are still more than 2000 Articles. Program stopped to avoid crashing system.")


#Move Files from temp back to /to_be_loaded
source = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\temp'
dest3 = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\_to be loaded/'

files = os.listdir(source)

for f in files:
        shutil.move(source+'\\'+f, dest3)

logging.info("Moved any offending Load Files from /temp back to to_be_loaded")


#Move Files to Completed New Folder
try:
    xlApp.Run('SaveAs')

except:
    xlApp.Quit()
    del xlApp
    logging.info("SaveAs Ran")
    date = datetime.today().strftime('%Y-%m-%d')
    dirname = completed+date+(" #1")
    if not os.path.exists(dirname):
        os.mkdir(dirname)
        dest4 = dirname

        files = os.listdir(datasheetloading)

        logging.info("Completed Load Folder Made")

        for f in files:
            shutil.move(datasheetloading+'\\'+f, dest4)
            logging.info("Individual and Merged Load Files moved to Complete Folder")
        try:
            shutil.move(logdir,dirname)

        except:
            print ("Ignore this message")




        
    else:
        dirname2 = completed+date+(" #2")
        os.mkdir(dirname2)
        dest4 = dirname2

        files = os.listdir(datasheetloading)

        logging.info("Completed Load Folder Made")

        for f in files:
            shutil.move(datasheetloading+'\\'+f, dest4)
            logging.info("Individual and Merged Load Files moved to Complete Folder")
        try:
            shutil.move(logdir,dirname2)

        except:
            print ("Ignore this message")


# BOT Email
gmail_user = 'BOTdr.loader@gmail.com'  
gmail_password = 'Xswqaz7471'

sent_from = gmail_user  
to = ['joe.martin@rs-components.com']  
subject = 'Dr. Loader - Load Prepared'
body = ("A Load has been prepared for you. Please check the Log File in the Completed Load File folder for information")
message = 'Subject: {}\n\n{}'.format(subject, body)


try:  
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    server.sendmail(sent_from, to, message)
    server.close()
    logging.info("Email Sent to R Content Technical Inbox")

except:
    print ("Email Fail")
