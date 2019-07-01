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
logdir = r'M:\GlobalImageManagement\Datasheet Loading New\Doc-Loader-Bot1\MAIN\Log'

#Glob Dirs
subtobeloaded = r'//wfsrvgbco001003/Datasrv5/MPP/GlobalImageManagement/Datasheet Loading New/_to be loaded/*'
subdatasheetloading = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\Datasheet Loading/*'

#Remove Previous Log File
try:
        os.remove("Log")

except:
        removelogfile = ("Either there is no previous Log File to be deleted or the program has encountered an Error")

#Create Log File
logging.basicConfig(format='%(levelname)s: %(asctime)s: %(message)s', filename="Log", level=logging.INFO)
logging.warning(removelogfile)
logging.info("Log File Created")

#Move New Load Files into 'to_be_loaded'
try:
        files = os.listdir(datasheetloading)

        for f in files:
                shutil.move(datasheetloading+'\\'+f, tobeloaded)
        logging.info("Files in Datasheet Loading moved to To_be_loaded if any.")

except:
        logging.warning("Error: Could not move files from Datasheet Loading to To_be_loaded. Closing Program.")
        exit()
        
#Pull Files from 'to_be_loaded' - oldest files first
try:
        list_of_files = glob.glob(subtobeloaded)
        oldest_file = min(list_of_files, key=os.path.getctime)
        print (oldest_file)
        shutil.move(oldest_file, datasheetloading)
        logging.info("Oldest File in to_be_loaded folder has been moved to Datasheet Loading:")
        logging.info(oldest_file)

except:
        logging.warning("Error: Could not fetch oldest file from To_be_loaded. Closing Program.")
        exit()

# Create Loop
## used to allow the program to retry the process if there were over 2000 articles on the previous run
restart = 1
any_files = 0
while restart < 10:
        #Launch Excel and Execute Macros
        try:
                try:
                        wb.Close(False)
                        xlApp.Quit()
                        del xlApp
                        logging.info("Closed Previous COM Instance If on 2nd or above passthrough.")
                        xlApp = win32com.client.DispatchEx('Excel.Application')
                        xlsPath = os.path.expanduser('M:\GlobalImageManagement\Datasheet Loading New\Doc-Loader-Bot1\MAIN\Merge Spreadsheet.xlsm')
                        wb = xlApp.Workbooks.Open(Filename=xlsPath)
                        xlApp.Run('simpleXlsMerger')
                        logging.info("Merger Macro Ran")
                        xlApp.Run('Clean_Sort')
                        logging.info("Clean_Sort Macro Ran")
                        xlApp.Run('datavalidation')
                        logging.info("Article Number Check Macro Ran")
                except:
                        xlApp = win32com.client.DispatchEx('Excel.Application')
                        xlsPath = os.path.expanduser('M:\GlobalImageManagement\Datasheet Loading New\Doc-Loader-Bot1\MAIN\Merge Spreadsheet.xlsm')
                        wb = xlApp.Workbooks.Open(Filename=xlsPath)
                        xlApp.Run('simpleXlsMerger')
                        logging.info("Merger Macro Ran")
                        xlApp.Run('Clean_Sort')
                        logging.info("Clean_Sort Macro Ran")
                        xlApp.Run('datavalidation')
                        logging.info("Article Number Check Macro Ran")

        except:
                try:
                        wb.Close(False)
                        xlApp.Quit()
                        del xlApp
                        logging.warning("Error: Excel Failure. Terminating Program.")
                        exit()
                except:
                        logging.warning("Error: Excel Failure.")
                        exit()

        #Check for <2000
        f=open("number.txt","r", encoding="utf-16")
        number=(f.read())
        logging.info("Reading Number of Articles...")

        #Stop Loop if under Article Limit
        if int(number) <1002:
                if int(number) >999:
                    logging.info("Articles Found:")
                    logging.info(number)
                    restart = 11

                else:
                    if len(os.listdir(tobeloaded) ) == 0:
                        logging.info("All Possible Load files Loaded")
                        logging.info("Articles Found:")
                        logging.info(number)
                        restart = 11

                    else:
                        list_of_files = glob.glob(subtobeloaded)
                        oldest_file = min(list_of_files, key=os.path.getctime)
                        print (oldest_file)
                        shutil.move(oldest_file, datasheetloading)
                        logging.info("More Load Files to compile possible, gathered another Load File:")
                        logging.info(oldest_file)
            



        else:
                list_of_files = glob.glob(subdatasheetloading)
                latest_file = max(list_of_files, key=os.path.getctime)
                print (latest_file)
                shutil.move(latest_file, temp)
                logging.info("Over 1000 Article Limit, removed offending load file")

                restart = restart + 1
                if restart == 10:
                    logging.critical("Program has re-run 10 times, and there are still more than 2000 Articles. Program stopped to avoid crashing system.")


#Move Files from temp back to /to_be_loaded
try:
        source = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\temp'
        dest3 = r'\\wfsrvgbco001003\Datasrv5\MPP\GlobalImageManagement\Datasheet Loading New\_to be loaded/'

        files = os.listdir(source)

        for f in files:
                shutil.move(source+'\\'+f, dest3)

        logging.info("Moved any offending Load Files from /temp back to to_be_loaded")

except:
        logging.warning("Error: Could not move files from Temp to To_be_loaded. Closing Program.")
        exit()

#Move Files to Completed New Folder
try:
    xlApp.Run('SaveAs')

except:
    xlApp.Quit()
    del xlApp
    logging.info("SaveAs Macro Ran")
    logging.info("COM Memory removed")
    date = datetime.today().strftime('%Y-%m-%d')
    dirname = completed+date+(" #1")
    dirname2 = completed+date+(" #2")
    dirname3 = completed+date+(" #3")
    dirname4 = completed+date+(" #4")
    if not os.path.exists(dirname):
        os.mkdir(dirname)
        dest4 = dirname

        files = os.listdir(datasheetloading)

        logging.info("Completed Load Folder Made:")
        logging.info(dirname)

        for f in files:
            shutil.move(datasheetloading+'\\'+f, dest4)
            logging.info("Individual and Merged Load Files moved to Complete Folder")
        try:
            shutil.move(logdir,dirname)

        except:
            logging.info("Log File Moved to Completed Folder")




        
    elif not os.path.exists(dirname2):
        os.mkdir(dirname2)
        dest4 = dirname2

        files = os.listdir(datasheetloading)

        logging.info("Completed Load Folder Made")
        logging.info(dirname2)

        for f in files:
            shutil.move(datasheetloading+'\\'+f, dest4)
            logging.info("Individual and Merged Load Files moved to Complete Folder")
        try:
            shutil.move(logdir,dirname2)

        except:
            logging.info("Log File Moved to Completed Folder")


    elif not os.path.exists(dirname3):
        os.mkdir(dirname3)
        dest4 = dirname3

        files = os.listdir(datasheetloading)

        logging.info("Completed Load Folder Made")
        logging.info(dirname3)

        for f in files:
            shutil.move(datasheetloading+'\\'+f, dest4)
            logging.info("Individual and Merged Load Files moved to Complete Folder")
        try:
            shutil.move(logdir,dirname3)

        except:
            logging.info("Log File Moved to Completed Folder")


    elif not os.path.exists(dirname4):
        os.mkdir(dirname4)
        dest4 = dirname4

        files = os.listdir(datasheetloading)

        logging.info("Completed Load Folder Made")
        logging.info(dirname4)

        for f in files:
            shutil.move(datasheetloading+'\\'+f, dest4)
            logging.info("Individual and Merged Load Files moved to Complete Folder")
        try:
            shutil.move(logdir,dirname4)

        except:
            logging.info("Log File Moved to Completed Folder")

logging.info("Program Ended Successfully")
