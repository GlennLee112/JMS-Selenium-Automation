from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime, timedelta, date

import dateutil.relativedelta as datereldel
from selenium import webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
import os
import shutil
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException

# Tkinter
import tkinter as tk
import tkinter.messagebox as msgbox

# Unused module; backup
# import time
# import os

# Unused function (backup)
def round_nearest_second(dt):
    if dt.microsecond >=400000:
        dt+=datereldel.relativedelta(seconds=1)
    return dt.replace(microsecond=0)

def showMessage(message, timeout=5000):
    root = tk.Tk()
    root.withdraw()
    root.after(timeout, root.destroy)
    msgbox.showinfo('Info', message, master=root)

# JMS Part
def Refresh_JMS(driver):
    driver.get('https://jms.jtexpress.my/index')
def Launch_JMS():
    # 1.1 Download latest chrome driver
    lcd=chromedriver_autoinstaller.install(cwd=True)

    # 1.2 Move downloaded chrome driver to root folder
    cdn = os.path.basename(lcd) # cdn = chrome driver name - require for renaming driver
    ndl = os.path.join(os.getcwd(),cdn) # ndl = new driver location - cwd path with chrome driver name

    # move to cwd with name
    shutil.move(lcd,ndl)

    # delete old driver folder
    shutil.rmtree(os.path.dirname(lcd))

    # 1.3 setting option
    DownLoc_Def=os.path.join(os.getcwd(),"Download")

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("user-agent=Chrome/80.0.3987.132")
    chrome_options.add_argument('user-data-dir=C:\\selenium\\ChromeProfile')
    chrome_options.add_experimental_option("detach",True)
    # Detach solution obtain from: https://stackoverflow.com/questions/74869011/reconnecting-to-preexisting-browser-in-selenium-causes-no-connection-could-be-m
    prefs={"download.default_directory":DownLoc_Def}
    chrome_options.add_experimental_option("prefs",prefs)
    #chrome_driver = "C:\\selenium\\chromedriver.exe" # Unncessary; will read ChromeDriver in the root folder automatically

    driver = webdriver.Chrome(options=chrome_options)

    # 1.4 Launching JMS website

    driver.get('https://jms.jtexpress.my/index')

    # Handle in case already log in case by finding 'Parcel Tracking' text / button
    try:
        #driver.find_element(By.XPATH, "//*[contains(text(), 'Parcel Tracking')]")
        WebDriverWait(driver,2).until(EC.presence_of_all_elements_located((By.XPATH, "//*[contains(text(),'页面长时间未操作，请重新登录系统')]")))
        print("You've been logged out due to inactivity, please log in again")
        time.sleep(2)
        pass
    except:
        pass
    # Detect if already logged in
    try:
        WebDriverWait(driver,2).until(EC.none_of(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Please Enter the Password']")),
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Please Enter The Employee ID']"))
        ))
        print("Already logged in")
    except:
        pass
    # Detect log in page in case not logged in; keep looping until user logged in / log in elements no longer present
    try:
        WebDriverWait(driver,2).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Please Enter The Employee ID']")))
        print("Please log in")
        pass
    except:
        pass
    try:
        WebDriverWait(driver,600).until(EC.none_of(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Please Enter the Password']")),
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Please Enter The Employee ID']"))
        ))
    except TimeoutException as TOE:
        print(str(TOE)+": You did not log in in time, please relaunch the 'Run' program")

def Driver():
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    driver = webdriver.Chrome(options=chrome_options)

    return driver
# Date_Range creator

def Date_range(start,end):
    start = datetime.strptime(start,"%Y-%m-%d").date()
    end = datetime.strptime(end,"%Y-%m-%d").date()

    # Set up variable to be used:
    ## Add one day (timedelta of 1)    
    one_data_delta = timedelta(days=1)
    
    ## Create list, including the first day
    date_list = [start]

    while end != start:     # While start day not equal to end day
        start += one_data_delta # Add one day into start day in a while loop
        date_list.append(start) # Append next day into date_list
        
    return date_list

def Delivery_Details(driver, Date, Old = "N"):
    # Delivery monitoring: -> open -> detail tab -> insert date(s) # Final Version V1 complete

    # 1. Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory

    # Download Path logic gate
    ## Normal (to standard folder)
    if Old == "N":
        dp_dm = os.path.join(os.getcwd(), "Delivery Monitoring")
    ## Historial/Old (to 'old' folder)
    else:
        dp_dm = os.path.join(os.getcwd(), "Delivery Monitoring", "Old")

    driver.command_executor._commands['send_command'] = (
        'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    params = {
        'cmd': 'Page.setDownloadBehavior',
        'params': {'behavior': 'allow', 'downloadPath': dp_dm}
    }
    driver.execute("send_command", params)

    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)
    # 2.1 Search box type in

    JMSSearch = driver.find_element(By.XPATH, "//input[@placeholder='Please check']")
    JMSSearch.send_keys("Delivery Monitoring")
    time.sleep(1)
    JMSSearch.send_keys(Keys.DOWN)
    JMSSearch.send_keys(Keys.ENTER)

    # 3.1 Switch to details tab
    DT = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//div[@id='tab-second']")))
    DT.click()

    time.sleep(1)
    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1 = 0
    tbi2 = 1
    for d in Date:
        d = str(d)
        # Click 'Start time' date box

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,
                                                                        "//label[contains(text(), 'Start Time:')]/following-sibling::div//input[@placeholder='Choose date and time']")))
        stb = driver.find_elements(By.XPATH,
                                   "//label[contains(text(), 'Start Time:')]/following-sibling::div//input[@placeholder='Choose date and time']")
        # driver.execute_script("arguments[0].click();",stb[0])
        stb[0].click()

        # Insert first date
        timebox = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))

        timebox[tbi1].send_keys(Keys.CONTROL + 'a')
        timebox[tbi1].send_keys(Keys.DELETE)
        timebox[tbi1].send_keys(d)
        #time.sleep(0.2)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi1 < 1:
            tbi1 = tbi1 + 1
        else:
            pass
        # Click 'End Time' date box
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,
                                                                        "//label[contains(text(), 'End Time:')]/following-sibling::div//input[@placeholder='Choose date and time']")))
        etb = driver.find_elements(By.XPATH,
                                   "//label[contains(text(), 'End Time:')]/following-sibling::div//input[@placeholder='Choose date and time']")
        etb[0].click()
        # Insert second date
        timebox = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        #time.sleep(0.1)
        # tbi2=1 # tbi = time box index
        timebox[tbi2].send_keys(Keys.CONTROL + 'a')
        timebox[tbi2].send_keys(Keys.DELETE)
        timebox[tbi2].send_keys(d)
        #time.sleep(0.2)

        # Note: end time box index stays at '1'; no if condition needed for changing index
        # Click Inquiry button
        Inq = driver.find_elements(By.XPATH,
                                   "//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
        # Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
        Inq[1].click()
        # Click Export button
        Exp = driver.find_elements(By.XPATH,
                                   "//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
        Exp[(len(Exp)) - 1].click()
        # Click Export pop up box 'Yes'
        Expp = driver.find_element(By.XPATH,
                                   "//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
        # Logging time when Export button is clicked
        Expp_T = datetime.now()
        # Click 'Download' after exporting
        Expp.click()  # Return the last index no. of found button and select
        # Click on Download Center and open up Download Center pop up
        time.sleep(0.2)
        DownC = driver.find_elements(By.XPATH,
                                     "//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
        DownC[(len(DownC)) - 1].click()
        EFD = driver.find_elements(By.XPATH, "//*[@aria-label='Export File Download']")

        # Needed Time and Module Name info for finding file to download in DC pop up.
        ## Date 1 (original)
        Expp_T1 = Expp_T.strftime('%Y-%m-%d %H:%M:%S')

        ## Date 2 (with one second margin)
        Expp_T2 = Expp_T + timedelta(0, 1)
        Expp_T2 = Expp_T2.strftime('%Y-%m-%d %H:%M:%S')

        # Module name
        module_name = '派件监控Details'

        # Expp_T time log as conditon for row to look for
        # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")

        # XPath for file with Completed status
        ## XPATH with original time
        SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        ## XPATH with alternative time (1 second more)
        SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"

        # Loop for 'Completed' status; repeat if no 'Completed' status

        while len(EFD[1].find_elements(By.XPATH, SFXPATH1)) == 0 and len(
                EFD[1].find_elements(By.XPATH, SFXPATH2)) == 0:  # When both time as not found
            # IB = inquiry button syntax
            IB = EFD[1].find_elements(By.XPATH, "//button[contains(., 'Inquiry')]")
            IB[len(IB) - 1].click()  # Set to the last

        # Condition between two path: one with original time, the other with one second margin or error.
        if len(EFD[1].find_elements(By.XPATH, SFXPATH1)) > 0:  # Original Time
            SFR = EFD[1].find_elements(By.XPATH, SFXPATH1)
        else:
            SFR = EFD[1].find_elements(By.XPATH, SFXPATH2)  # Alternative time (1 second more)

        SFRD = SFR[0]  # Search File Read

        # DB = Download Button
        DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
        DB.click()  # Click Download Button

        Close = EFD[1].find_elements(By.XPATH,
                                     "//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
        Close[1].click()

        # Wait file to complete download using while loop
        ## Download Folder
        dl_wait = True  # Default True
        while dl_wait:
            time.sleep(1)

            dl_wait = False
            for fname in os.listdir(dp_dm):
                if fname.endswith('.crdownload'):
                    dl_wait = True

    # Close Delivery Monitoring > Download Center
    # DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    # DMDC_Close[len(DMDC_Close)-1].click()

    time.sleep(0.2)

    # Close Delivery Monitoring tab
    DM_Close = driver.find_elements(By.XPATH,
                                    "//div[@id='tab-DistributeMonitor|crisbiIndex' and contains(text(), 'Delivery Monitoring')]//span[contains(@class, 'el-icon-close')]")
    DM_Close[len(DM_Close) - 1].click()

# Vocabulary

## DMO = Delivery Monitoring: - Open
## DMDD = Delivery monitoring: - detail tab - date
## DMODD = Delivery monitoring: - open - detail tab - insert date(s)

# AMODD 
def Arrival(driver,Date,Old='N'): # Arrival monitoring -> open -> date(s) # Final Version V1 complete

    # 1. Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory
    
    driver.command_executor._commands['send_command'] = (
    'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    if Old=='N':
        dp_am = os.path.join(os.getcwd(),"Arrival Monitoring") # dp_dm = download path - delivery monitoring
    elif Old!='N':
        dp_am = os.path.join(os.getcwd(),"Arrival Monitoring","Old")
    params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': { 'behavior': 'allow', 'downloadPath': dp_am }
    }
    driver.execute("send_command", params)
   
    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)

    # 2.1 Search box type in

    JMSSearch=driver.find_element(By.XPATH,"//input[@placeholder='Please check']")
    JMSSearch.send_keys("Arrival Monitoring")
    time.sleep(1)
    JMSSearch.send_keys(Keys.DOWN)
    JMSSearch.send_keys(Keys.ENTER)

    # 2.2 Select Drop point
    # Click type drop list
    Select_type = WebDriverWait(driver, 60).until(EC.presence_of_element_located(
        (By.XPATH,
         "//span[@title='Type :']/ancestor::div[@class='el-form-item el-form-item--small']//div[contains(@class, 'el-input el-input--small el-input--suffix')]")))
    Select_type.click()
    # Click drop point
    Drop_Point = WebDriverWait(driver, 60).until(EC.presence_of_element_located(
        (By.XPATH, "//div[contains(@class, 'el-select-dropdown el-popper')]//span[text()='Drop Point']")))
    try:
        Drop_Point.click()
    except ElementNotInteractableException:
        pass

    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1=0 # For start time
    tbi2=0 # For end time
    
    for d in Date:
        d=str(d)
        # Click 'Start time' date box
        stb=etb=driver.find_elements(By.XPATH, "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")
        stb[0].click()
        
        # Insert first date
        timebox=WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #print(tbi1)  print tbi1 for debugging purpose; remove '#' toe enable
        
        timebox[tbi1].send_keys(Keys.CONTROL + 'a')
        timebox[tbi1].send_keys(Keys.DELETE)
        timebox[tbi1].send_keys(d)
        #timebox[tbi1].send_keys(Keys.ENTER)
        #timebox[tbi1].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi1 < 1:
            tbi1+=1
        else:
            tbi1
        
        # Click 'End Time' date box
        etb[1].click()
        # Insert second date
        timebox=WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #tbi2=1 # tbi = time box index
        timebox[tbi2].send_keys(Keys.CONTROL + 'a')
        timebox[tbi2].send_keys(Keys.DELETE)
        timebox[tbi2].send_keys(d)
        #timebox[tbi2].send_keys(Keys.ENTER)
        #timebox[tbi2].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi2 < 1:
            tbi2+=1
        else:
            tbi2

        # Click Inquiry button
        Inq=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
        #Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
        Inq[0].click()
        # Click Export button
        Exp=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
        Exp[0].click()
        
        # Click Export pop up box 'Yes'
        Expp=driver.find_element(By.XPATH,"//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
        # Logging time when Export button is clicked
        Expp_T=datetime.now()
        # Click 'Download' after exporting
        Expp.click() # Return the last index no. of found button and select
        # Click on Download Center and open up Download Center pop up
        time.sleep(0.2)
        DownC=driver.find_elements(By.XPATH,"//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
        DownC[0].click()
        EFD = driver.find_element(By.XPATH,"//*[@aria-label='Export File Download']")
        
        # Needed Time and Module Name info for finding file to download in DC pop up.
        ## Date 1 (original)
        Expp_T1=Expp_T.strftime('%Y-%m-%d %H:%M:%S')
        
        ## Date 2 (with one second margin)
        Expp_T2=Expp_T+timedelta(0,1)
        Expp_T2=Expp_T2.strftime('%Y-%m-%d %H:%M:%S')
        
        # Module name
        module_name = '到件监控Summary'
        
        # Expp_T time log as conditon for row to look for
        # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")
        
        # XPath for file with Completed status
        ## Search File XPATH with original time
        SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        ## XPATH with alternative time (1 second more)
        SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        
        # Loop for 'Completed' status; repeat if no 'Completed' status
        
        while len(EFD.find_elements(By.XPATH, SFXPATH1)) == 0 and len(EFD.find_elements(By.XPATH, SFXPATH2)) == 0: # When both time as not found
                # IB = inquiry button syntax
                EFD_IB=EFD.find_elements(By.XPATH,"//button[contains(., 'Inquiry')]")
                EFD_IB[1].click()

        # Condition between two path: one with original time, the other with one second margin or error.
        ## SFXPATH_Test for match case condition
        SFXPATH_Test = len(EFD.find_elements(By.XPATH, SFXPATH1))

        match SFXPATH_Test:
            # For case of without time delay - syntax with original time    
            case SFXPATH_Test if SFXPATH_Test > 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH1) # original time syntax
            # For case of one second late - syntax with one second margin added    
            case SFXPATH_Test if SFXPATH_Test == 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH2) 
            
        SFRD = SFR[0] # Search File Row to Download

        # DB = Download Button
        DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
        DB.click() # Click Download Button
        
        Close=EFD.find_element(By.XPATH, "//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
        Close.click()
        
        # Wait file to complete download using while loop
        ## Download Folder
        DownLoc_Def=os.path.join(os.getcwd(),"Arrival Monitoring")
        
        ## Hold until Chrome download temporary file no longer present
        dl_wait = True # Default True
        while dl_wait:
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(DownLoc_Def):
                if fname.endswith('.crdownload'):
                    dl_wait = True

    # Close Delivery Monitoring > Download Center
    #DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    #DMDC_Close[len(DMDC_Close)-1].click()
    
    time.sleep(0.2)
    
    # Close Delivery Monitoring tab
    AM_Close=driver.find_elements(By.XPATH,"//div[@id='tab-ArriveMonitor|crisbiIndex' and contains(text(), 'Arrival Monitoring (New)')]//span[contains(@class, 'el-icon-close')]")
    AM_Close[0].click()
    
# AMODD 
def Departure(driver,Date): # Departure monitoring (new) -> open -> Next Station button -> date(s) # Final Version V1 complete
    # 1. Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory
    
    driver.command_executor._commands['send_command'] = (
    'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    dp_dpm = os.path.join(os.getcwd(),"Departure Monitoring") # dp_dm = download path - delivery monitoring
    params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': { 'behavior': 'allow', 'downloadPath': dp_dpm }
    }
    driver.execute("send_command", params)
   
    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)
        # 2.1 Search box type in

    JMSSearch=driver.find_element(By.XPATH,"//input[@placeholder='Please check']")
    JMSSearch.send_keys("Departure Monitoring (New)")
    time.sleep(1)
    JMSSearch.send_keys(Keys.DOWN)
    JMSSearch.send_keys(Keys.ENTER)

    # Select Type
    # Select Type
    Select_type = WebDriverWait(driver, 60).until(EC.presence_of_element_located(
        (By.XPATH,
         "//span[@title='Type :']/ancestor::div[@class='el-form-item el-form-item--small']//div[contains(@class, 'el-input el-input--small el-input--suffix')]")))
    Select_type.click()
    # Click drop point
    Drop_Point = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
        (By.XPATH, "//div[contains(@class, 'el-select-dropdown el-popper')]//span[text()='Drop Point']")))
    Drop_Point = driver.find_elements(By.XPATH,
                                      "//div[contains(@class, 'el-select-dropdown el-popper')]//span[text()='Drop Point']")
    Drop_Point[len(Drop_Point) - 1].click()
    
    # Tick 'Previous Station'
    try:
        driver.find_element(By.XPATH,"//div[@class='el-form-item el-form-item--small']//span[@class='el-checkbox__input' and following-sibling::span[@class='el-checkbox__label' and text()='Next Station']]").click()
    except NoSuchElementException:
        pass

    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1=0 # For start time
    tbi2=1 # For end time
    
    for d in Date:
        d=str(d)
        # Click 'Start time' date box
        ## assign stb and etb to same element; element has the same syntax
        stb=etb=driver.find_elements(By.XPATH, "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")
        stb[0].click()
        
        # Insert first date
        timebox = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #print(tbi1)  print tbi1 for debugging purpose; remove '#' toe enable
        
        timebox[tbi1].send_keys(Keys.CONTROL + 'a')
        timebox[tbi1].send_keys(Keys.DELETE)
        timebox[tbi1].send_keys(d)
        #timebox[tbi1].send_keys(Keys.ENTER)
        #timebox[tbi1].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi1 < 1:
            tbi1+=1
        else:
            tbi1
        
        # Click 'End Time' date box
        etb[1].click()
        # Insert second date
        timebox = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #tbi2=1 # tbi = time box index
        timebox[tbi2].send_keys(Keys.CONTROL + 'a')
        timebox[tbi2].send_keys(Keys.DELETE)
        timebox[tbi2].send_keys(d)
        
        # Click Inquiry button
        Inq=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
        #Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
        Inq[0].click()
        # Click Export button
        Exp=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
        Exp[0].click()
        
        # Click Export pop up box 'Yes'
        Expp=driver.find_element(By.XPATH,"//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
        # Logging time when Export button is clicked
        Expp_T=datetime.now()
        # Click 'Download' after exporting
        Expp.click() # Return the last index no. of found button and select
        # Click on Download Center and open up Download Center pop up
        time.sleep(0.2)
        DownC=driver.find_elements(By.XPATH,"//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
        DownC[0].click()
        EFD = driver.find_element(By.XPATH,"//*[@aria-label='Export File Download']")
        
        # Needed Time and Module Name info for finding file to download in DC pop up.
        ## Date 1 (original)
        Expp_T1=Expp_T.strftime('%Y-%m-%d %H:%M:%S')
        
        ## Date 2 (with one second margin)
        Expp_T2=Expp_T+timedelta(0,1)
        Expp_T2=Expp_T2.strftime('%Y-%m-%d %H:%M:%S')
        
        # Module name
        module_name = '发件监控Summary'
        #print(Expp_T1)
        
        # Expp_T time log as conditon for row to look for
        # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")
        
        # XPath for file with Completed status
        ## Search File XPATH with original time
        SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        ## XPATH with alternative time (1 second more)
        SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        
        # Loop for 'Completed' status; repeat if no 'Completed' status
        
        while len(EFD.find_elements(By.XPATH, SFXPATH1)) == 0 and len(EFD.find_elements(By.XPATH, SFXPATH2)) == 0: # When both time as not found
                # IB = inquiry button syntax
                EFD_IB=EFD.find_elements(By.XPATH,"//button[contains(., 'Inquiry')]")
                EFD_IB[1].click()

        # Condition between two path: one with original time, the other with one second margin or error.
        ## SFXPATH_Test for match case condition
        SFXPATH_Test = len(EFD.find_elements(By.XPATH, SFXPATH1))

        match SFXPATH_Test:
            # For case of without time delay - syntax with original time    
            case SFXPATH_Test if SFXPATH_Test > 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH1) # original time syntax
            # For case of one second late - syntax with one second margin added    
            case SFXPATH_Test if SFXPATH_Test == 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH2) 
            
        SFRD = SFR[0] # Search File Row to Download

        # DB = Download Button
        DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
        DB.click() # Click Download Button
        
        Close=EFD.find_element(By.XPATH, "//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
        Close.click()
        
        # Wait file to complete download using while loop
        ## Download Folder
        #DownLoc_Def=os.path.join(os.getcwd(),"Departure Monitoring")
        
        ## Hold until Chrome download temporary file no longer present
        dl_wait = True # Default True
        while dl_wait:
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(dp_dpm):
                if fname.endswith('.crdownload'):
                    dl_wait = True

    # Close Delivery Monitoring > Download Center
    #DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    #DMDC_Close[len(DMDC_Close)-1].click()
    
    time.sleep(0.2)
    
    # Close Delivery Monitoring tab
    DPM_Close=driver.find_elements(By.XPATH,"//div[@id='tab-SendOutMonitor|crisbiIndex' and contains(text(), 'Departure Monitoring (New)')]//span[contains(@class, 'el-icon-close')]")
    DPM_Close[0].click()

## Overnight 
def Overnight(driver,Date=""):
    # 1. Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory

    driver.command_executor._commands['send_command'] = (
    'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    dp_gdpo = os.path.join(os.getcwd(),"General DP Overnight") # dp_dm = download path - delivery monitoring
    params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': { 'behavior': 'allow', 'downloadPath': dp_gdpo }
    }
    driver.execute("send_command", params)

    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)

    # 2.1 Search box type in 'General DP overnight monitoring'

    JMSSearch=driver.find_element(By.XPATH,"//input[@placeholder='Please check']")
    JMSSearch.send_keys("General DP Overnight Monitoring")
    time.sleep(1)
    JMSSearch.send_keys(Keys.DOWN)
    JMSSearch.send_keys(Keys.ENTER)

    # 3.1 Switch to details tab
    DT=WebDriverWait(driver,50).until(EC.presence_of_element_located((By.XPATH,"//div[@id='tab-detail']")))
    DT.click()
    time.sleep(1)

    # 3.2 Populate start and end time box for general dp data

    ## Date variable logic gate
    ### if empty, default to today's date (i.e.: end date of yesterday; start date of previous month of yesterday)

    if Date == "":

        Now = datetime.now()
        ## Start Date
        End_Date=Now-datereldel.relativedelta(days=1)
        End_Date=End_Date.strftime('%Y-%m-%d')

        ## End Date
        Start_Date=Now-datereldel.relativedelta(months=1)
        Start_Date=Start_Date.strftime('%Y-%m-%d')

    ### If date variable is not left empty and is a list (as it should), assign min and max date of the list as start and end date, respectively.
    ### Isinstance reference: https://stackoverflow.com/questions/26544091/checking-if-type-list-in-python

    elif Date != "" and isinstance(Date, list) == True:

        Start_Date=min(Date)
        Start_Date=Start_Date.strftime('%Y-%m-%d')
        
        End_Date=max(Date)
        End_Date=End_Date.strftime('%Y-%m-%d')

    elif Date != '' and isinstance(Date, list) == False and (isinstance(Date,date) == True or isinstance(Date,datetime) == True or isinstance(Date,str) == True) :
        if isinstance(Date,datetime) == True:
            # Datetime -- convert to date
            Date=Date.date()
        elif isinstance(Date,str) == True and isinstance(Date,list) == False:
            # Text -- convert to date
            Date=datetime.strptime(Date,"%Y-%m-%d").date()
        else:
            pass
        
        Now = Date
        ## End Date
        End_Date=Date-datereldel.relativedelta(days=1)
        End_Date=End_Date.strftime('%Y-%m-%d')

        ## Start Date
        Start_Date=(Date-datereldel.relativedelta(days=1))-datereldel.relativedelta(months=1)
        Start_Date=Start_Date.strftime('%Y-%m-%d')
        
        
    else:
        raise TypeError('You Date must be in Panda datetime or date format, as a single date or a pair list of start and end date; please amend.')

    ## 3.2.1 Insert Start Date
    stb=driver.find_elements(By.XPATH, "//label[@for='startTime']/following-sibling::div//input[@placeholder='Select Date']")
    #stb=driver.find_elements(By.XPATH, "//label[contains(text(), 'Start Date:')]/following-sibling::div//input[@placeholder='Select Date']")
    #stb=stb[len(stb)-len(stb)]
    stb=stb[0]
    stb.send_keys(Keys.CONTROL + 'a')
    stb.send_keys(Keys.DELETE)
    stb.send_keys(Start_Date)

    ## 3.2.1 Insert End Date
    etb=driver.find_elements(By.XPATH, "//label[@for='endTime']/following-sibling::div//input[@placeholder='Select Date']")
    #etb=etb[len(etb)-len(etb)]
    etb=etb[0]
    etb.send_keys(Keys.CONTROL + 'a')
    etb.send_keys(Keys.DELETE)
    etb.send_keys(End_Date)

    # Click Inquiry button
    Inq=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
    Inq[1].click()

    # Click Export button
    Exp=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
    Expp_T=datetime.now() # Log time when Expp button is pressed
    Exp[1].click()

    # Click on Download Center and open find the pop-up Export File Download
    time.sleep(0.2)
    DownC=driver.find_elements(By.XPATH,"//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
    DownC[1].click()
    EFD = driver.find_element(By.XPATH,"//*[@aria-label='Export File Download']")

    # Needed Time and Module Name info for finding file to download in DC pop up.
    ## Date 1 (original)
    Expp_T1=Expp_T.strftime('%Y-%m-%d %H:%M:%S')

    ## Date 2 (with one second margin)
    Expp_T2=Expp_T+timedelta(0,1)
    Expp_T2=Expp_T2.strftime('%Y-%m-%d %H:%M:%S')

    # Module name
    module_name = 'General DP Overnight Monitoring'

    # Expp_T time log as conditon for row to look for
    # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")

    # XPath for file with Completed status (OG time and offset time)
    ## Search File XPATH with OG time
    SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
    ## Search File XPATH with offset time
    SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"

    # Loop for 'Completed' status; repeat if no 'Completed' status and click 'Inqury' button
    while len(EFD.find_elements(By.XPATH, SFXPATH1)) == 0 and len(EFD.find_elements(By.XPATH, SFXPATH2)) == 0: # When both time as not found
            # IB = inquiry button syntax
            EFD_IB=EFD.find_elements(By.XPATH,"//button[contains(., 'Inquiry')]")
            EFD_IB[len(EFD_IB)-1].click()

    # Condition between two path: one with original time, the other with one second margin or error.
    ## SFXPATH_Test for match case condition
    SFXPATH_Test = len(EFD.find_elements(By.XPATH, SFXPATH1))

    match SFXPATH_Test:
        # For case of without time delay - syntax with original time    
        case SFXPATH_Test if SFXPATH_Test > 0:
            SFR = EFD.find_elements(By.XPATH, SFXPATH1) # original time syntax
        # For case of one second late - syntax with one second margin added    
        case SFXPATH_Test if SFXPATH_Test == 0:
            SFR = EFD.find_elements(By.XPATH, SFXPATH2)

    SFRD = SFR[0] # Search File Row to Download

    # DB = Download Button
    DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
    DB.click() # Click Download Button

    Close=EFD.find_elements(By.XPATH, "//div[@aria-label='Export File Download']//button[@aria-label='Close']")
    Close[1].click()

    # Wait file to complete download using while loop
    ## Download Folder
    #DownLoc_Def=os.path.join(os.getcwd(),"Departure Monitoring")

    ## Hold until Chrome download temporary file no longer present
    dl_wait = True # Default True
    while dl_wait:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(dp_gdpo):
            if fname.endswith('.crdownload'):
                dl_wait = True

    # Close Delivery Monitoring > Download Center
    #DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    #DMDC_Close[len(DMDC_Close)-1].click()

    time.sleep(0.2)

    # Close Delivery Monitoring tab
    GDOM_Close=driver.find_elements(By.XPATH,"//div[@id='tab-WarehouseRetentionMonitor' and contains(text(), 'General DP Overnight')]//span[contains(@class, 'el-icon-close')]")
    GDOM_Close[0].click()
    
def Delivery_Summary(driver,Date,DownLoc="",Type=""):
   # 1. Reconnect Chrome

    # 1.Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory
    
    driver.command_executor._commands['send_command'] = (
    'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    if DownLoc=="":
        dp_dms = os.path.join(os.getcwd(),"Delivery Monitoring Summary") # dp_dm = download path - delivery monitoring
    else:
        dp_dms=DownLoc
    params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': { 'behavior': 'allow', 'downloadPath': dp_dms }
    }
    driver.execute("send_command", params)
   
    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)
        # 2.1 Search box type in

    JMSSearch=driver.find_element(By.XPATH,"//input[@placeholder='Please check']")
    JMSSearch.send_keys("Delivery Monitoring")
    time.sleep(1)
    AutoFill=driver.find_element(By.XPATH, "//ul[@class='el-scrollbar__view el-autocomplete-suggestion__list' and @role='listbox']//li[@role='option' and normalize-space(text())='Delivery Monitoring']")
    AutoFill.click()

    # Select type: Drop point
    ## Click drop down box
    if Type=="" or Type =='1':
        Type="Drop Point"
    elif Type == '2':
        Type="Region"
    elif Type == '3':
        Type="Delivery Dispatcher"
    
    Select_type=driver.find_elements(By.XPATH,"//span[@title='Type :']/ancestor::div[@class='el-form-item el-form-item--small']//div[contains(@class, 'el-input el-input--small el-input--suffix')]")
    Select_type[len(Select_type)-1].click()
    ## Select 'Drop point'
    Droppoint=driver.find_elements(By.XPATH,f"//ul[@class='el-scrollbar__view el-select-dropdown__list']//li[span='{Type}']")
    time.sleep(1)
    Droppoint[0].click()
    
    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1=0
    tbi2=1
    
    for d in Date:
        d=str(d)
        # Click 'Start time' date box (first element)
        tb = driver.find_elements(By.XPATH, "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")
        tb[0].click()
        
        # Insert first date
        timebox=WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.2)
        timebox[tbi1].send_keys(Keys.CONTROL + 'a')
        timebox[tbi1].send_keys(Keys.DELETE)
        timebox[tbi1].send_keys(d)
        time.sleep(0.1)        
        # If logic gate to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi1 < 1:
            tbi1+=1
        else:
            tbi1

        time.sleep(0.1)
        # Click 'End Time' date box
        #etb=driver.find_elements(By.XPATH, "//label[contains(text(), 'End Time:')]/following-sibling::div//input[@placeholder='Choose date and time']")
        tb[1].click()
        # Insert second date
        timebox=WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #tbi2=1 # tbi = time box index
        #print(tbi2)
        timebox[tbi2].send_keys(Keys.CONTROL + 'a')
        timebox[tbi2].send_keys(Keys.DELETE)
        timebox[tbi2].send_keys(d)
        time.sleep(0.2)
        # Note: end time box index stays at '1'; no if condition needed for changing index
        
        # Click Inquiry button
        Inq=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
        #Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
        Inq[0].click()
        # Click Export button
        Exp=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
        Exp[0].click()
        # Click Export pop up box 'Yes'
        Expp=driver.find_element(By.XPATH,"//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
        # Logging time when Export button is clicked
        Expp_T=datetime.now()
        # Click 'Download' after exporting
        Expp.click() # Return the last index no. of found button and select
        # Click on Download Center and open up Download Center pop up
        time.sleep(0.2)
        DownC=driver.find_elements(By.XPATH,"//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
        DownC[0].click()
        EFD = driver.find_elements(By.XPATH,"//*[@aria-label='Export File Download']")
        
        # Needed Time and Module Name info for finding file to download in DC pop up.
        ## Date 1 (original)
        Expp_T1=Expp_T.strftime('%Y-%m-%d %H:%M:%S')
        
        ## Date 2 (with one second margin)
        Expp_T2=Expp_T+timedelta(0,1)
        Expp_T2=Expp_T2.strftime('%Y-%m-%d %H:%M:%S')
        
        # Module name
        module_name = '派件监控Summary'
        
        # Expp_T time log as conditon for row to look for
        # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")
        
        # XPath for file with Completed status
        ## XPATH with original time
        SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        ## XPATH with alternative time (1 second more)
        SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        
        # Loop for 'Completed' status; repeat if no 'Completed' status
        
        while len(EFD[1].find_elements(By.XPATH, SFXPATH1)) == 0 and len(EFD[1].find_elements(By.XPATH, SFXPATH2)) == 0: # When both time as not found
                # IB = inquiry button syntax
                IB=EFD[1].find_elements(By.XPATH, "//button[contains(., 'Inquiry')]")
                IB[1].click() # Set to the last

        # Condition between two path: one with original time, the other with one second margin or error.
        if len(EFD[1].find_elements(By.XPATH, SFXPATH1)) > 0: # Original Time
            SFR = EFD[1].find_elements(By.XPATH, SFXPATH1)
        else:
            SFR = EFD[1].find_elements(By.XPATH, SFXPATH2)    # Alternative time (1 second more)
            
        SFRD = SFR[0] # Search File Read

        # DB = Download Button
        DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
        DB.click() # Click Download Button
        
        Close=EFD[1].find_elements(By.XPATH, "//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
        Close[0].click()
        
        # Wait file to complete download using while loop
        ## Download Folder
        ##DownLoc_Def=os.path.join(os.getcwd(),"Download")
        
        ## Hold until Chrome download temporary file no longer present

        # Reference: https://stackoverflow.com/questions/63637077/how-to-wait-for-a-file-to-be-downloaded-in-selenium-and-python-before-moving-for

        dl_wait = True # Default True
        while dl_wait:
            #
            dp_dms = os.path.join(os.getcwd(), "Delivery Monitoring Summary")
            # Loop
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(dp_dms):
                if fname.endswith('.crdownload'):
                    dl_wait = True

    # Close Delivery Monitoring > Download Center
    #DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    #DMDC_Close[len(DMDC_Close)-1].click()
    
    time.sleep(0.2)
    
    # Close Delivery Monitoring tab
    DM_Close=driver.find_elements(By.XPATH,"//div[@id='tab-DistributeMonitor|crisbiIndex' and contains(text(), 'Delivery Monitoring')]//span[contains(@class, 'el-icon-close')]")
    DM_Close[len(DM_Close)-1].click()

def Delivery_Summary_Lump(driver, Date_1, Date_2, DownLoc="", Type=""):
    # 1. Reconnect Chrome

    # 1.Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory

    driver.command_executor._commands['send_command'] = (
        'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    if DownLoc == "":
        dp_dms = os.path.join(os.getcwd(), "Delivery Monitoring Summary (Lump)")  # dp_dm = download path - delivery monitoring
    else:
        dp_dms = DownLoc

    params = {
        'cmd': 'Page.setDownloadBehavior',
        'params': {'behavior': 'allow', 'downloadPath': dp_dms}
    }
    driver.execute("send_command", params)

    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)
    # 2.1 Search box type in

    JMSSearch = driver.find_element(By.XPATH, "//input[@placeholder='Please check']")
    JMSSearch.send_keys("Delivery Monitoring")
    time.sleep(1)
    AutoFill = driver.find_element(By.XPATH,
                                   "//ul[@class='el-scrollbar__view el-autocomplete-suggestion__list' and @role='listbox']//li[@role='option' and normalize-space(text())='Delivery Monitoring']")
    AutoFill.click()

    # Select type: Drop point
    ## Click drop down box
    if Type == "" or Type == '1':
        Type = "Drop Point"
    elif Type == '2':
        Type = "Region"
    elif Type == '3':
        Type = "Delivery Dispatcher"

    Select_type = driver.find_elements(By.XPATH,
                                       "//span[@title='Type :']/ancestor::div[@class='el-form-item el-form-item--small']//div[contains(@class, 'el-input el-input--small el-input--suffix')]")
    Select_type[len(Select_type) - 1].click()
    ## Select 'Drop point'
    Droppoint = driver.find_elements(By.XPATH,
                                     f"//ul[@class='el-scrollbar__view el-select-dropdown__list']//li[span='{Type}']")
    time.sleep(1)
    Droppoint[0].click()

    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1 = 0
    tbi2 = 1

    #for d in Date:
    d1 = str(Date_1)
    # Click 'Start time' date box (first element)
    tb = driver.find_elements(By.XPATH,
                              "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")
    tb[0].click()

    # Insert first date
    timebox = WebDriverWait(driver, 60).until(
        EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
    time.sleep(0.2)
    timebox[tbi1].send_keys(Keys.CONTROL + 'a')
    timebox[tbi1].send_keys(Keys.DELETE)
    timebox[tbi1].send_keys(d1)
    time.sleep(0.1)

    # Click 'End Time' date box
    # etb=driver.find_elements(By.XPATH, "//label[contains(text(), 'End Time:')]/following-sibling::div//input[@placeholder='Choose date and time']")
    d2 = str(Date_2)
    tb[1].click()
    # Insert second date
    timebox = WebDriverWait(driver, 60).until(
        EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
    time.sleep(0.1)
    # tbi2=1 # tbi = time box index
    # print(tbi2)
    timebox[tbi2].send_keys(Keys.CONTROL + 'a')
    timebox[tbi2].send_keys(Keys.DELETE)
    timebox[tbi2].send_keys(d2)
    time.sleep(0.2)
    # Note: end time box index stays at '1'; no if condition needed for changing index

    # Click Inquiry button
    Inq = driver.find_elements(By.XPATH,
                               "//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
    # Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
    Inq[0].click()
    # Click Export button
    Exp = driver.find_elements(By.XPATH,
                               "//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
    Exp[0].click()
    # Click Export pop up box 'Yes'
    Expp = driver.find_element(By.XPATH,
                               "//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
    # Logging time when Export button is clicked
    Expp_T = datetime.now()
    # Click 'Download' after exporting
    Expp.click()  # Return the last index no. of found button and select
    # Click on Download Center and open up Download Center pop up
    time.sleep(0.2)
    DownC = driver.find_elements(By.XPATH,
                                 "//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
    DownC[0].click()
    EFD = driver.find_elements(By.XPATH, "//*[@aria-label='Export File Download']")

    # Needed Time and Module Name info for finding file to download in DC pop up.
    ## Date 1 (original)
    Expp_T1 = Expp_T.strftime('%Y-%m-%d %H:%M:%S')

    ## Date 2 (with one second margin)
    Expp_T2 = Expp_T + timedelta(0, 1)
    Expp_T2 = Expp_T2.strftime('%Y-%m-%d %H:%M:%S')

    # Module name
    module_name = '派件监控Summary'

    # Expp_T time log as conditon for row to look for
    # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")

    # XPath for file with Completed status
    ## XPATH with original time
    SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
    ## XPATH with alternative time (1 second more)
    SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"

    # Loop for 'Completed' status; repeat if no 'Completed' status

    while len(EFD[1].find_elements(By.XPATH, SFXPATH1)) == 0 and len(
            EFD[1].find_elements(By.XPATH, SFXPATH2)) == 0:  # When both time as not found
        # IB = inquiry button syntax
        IB = EFD[1].find_elements(By.XPATH, "//button[contains(., 'Inquiry')]")
        IB[1].click()  # Set to the last

    # Condition between two path: one with original time, the other with one second margin or error.
    if len(EFD[1].find_elements(By.XPATH, SFXPATH1)) > 0:  # Original Time
        SFR = EFD[1].find_elements(By.XPATH, SFXPATH1)
    else:
        SFR = EFD[1].find_elements(By.XPATH, SFXPATH2)  # Alternative time (1 second more)

    SFRD = SFR[0]  # Search File Read

    # DB = Download Button
    DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
    DB.click()  # Click Download Button

    Close = EFD[1].find_elements(By.XPATH,
                                 "//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
    Close[0].click()

    # Wait file to complete download using while loop
    ## Download Folder
    ##DownLoc_Def=os.path.join(os.getcwd(),"Download")

    ## Hold until Chrome download temporary file no longer present

    # Reference: https://stackoverflow.com/questions/63637077/how-to-wait-for-a-file-to-be-downloaded-in-selenium-and-python-before-moving-for

    dl_wait = True  # Default True
    while dl_wait:
        #dp_dms = os.path.join(os.getcwd(), "Delivery Monitoring Summary")
        # Loop
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(dp_dms):
            if fname.endswith('.crdownload'):
                dl_wait = True

    # Close Delivery Monitoring > Download Center
    # DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    # DMDC_Close[len(DMDC_Close)-1].click()

    time.sleep(0.2)

    # Close Delivery Monitoring tab
    DM_Close = driver.find_elements(By.XPATH,
                                    "//div[@id='tab-DistributeMonitor|crisbiIndex' and contains(text(), 'Delivery Monitoring')]//span[contains(@class, 'el-icon-close')]")
    DM_Close[len(DM_Close) - 1].click()

def Arrival_Backup(driver,Date,Old='N'): # Arrival monitoring -> open -> date(s) # Final Version V1 complete

    # 1. Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory
    
    driver.command_executor._commands['send_command'] = (
    'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    if Old=='N':
        dp_am = os.path.join(os.getcwd(),"Arrival Monitoring") # dp_dm = download path - delivery monitoring
    elif Old!='N':
        dp_am = os.path.join(os.getcwd(),"Arrival Monitoring","Old")
    params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': { 'behavior': 'allow', 'downloadPath': dp_am }
    }
    driver.execute("send_command", params)
    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)
        # 2.1 Search box type in
    
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,"//input[@placeholder='Please check']")))
    JMSSearch=driver.find_element(By.XPATH,"//input[@placeholder='Please check']")
    JMSSearch.send_keys("Arrival Monitoring")
    time.sleep(1)
    JMSSearch.send_keys(Keys.DOWN)
    JMSSearch.send_keys(Keys.ENTER)
    

    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1=0 # For start time
    tbi2=0 # For end time
    
    for d in Date:
        d=str(d)
        # Click 'Start time' date box
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")))
        stb=etb=driver.find_elements(By.XPATH, "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")
        stb[0].click()
        
        # Insert first date
        timebox=WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #print(tbi1)  print tbi1 for debugging purpose; remove '#' toe enable
        
        timebox[tbi1].send_keys(Keys.CONTROL + 'a')
        timebox[tbi1].send_keys(Keys.DELETE)
        timebox[tbi1].send_keys(d)
        #timebox[tbi1].send_keys(Keys.ENTER)
        #timebox[tbi1].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi1 < 1:
            tbi1+=1
        else:
            tbi1
        
        # Click 'End Time' date box
        etb[1].click()
        # Insert second date
        timebox=WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        #tbi2=1 # tbi = time box index
        timebox[tbi2].send_keys(Keys.CONTROL + 'a')
        timebox[tbi2].send_keys(Keys.DELETE)
        timebox[tbi2].send_keys(d)
        #timebox[tbi2].send_keys(Keys.ENTER)
        #timebox[tbi2].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi2 < 1:
            tbi2+=1
        else:
            tbi2
        
        # Select type: Drop point
        # Click drop down box
        Select_type=WebDriverWait(driver,60).until(EC.presence_of_element_located((By.XPATH,"//span[@title='Type :']/ancestor::div[@class='el-form-item el-form-item--small']//div[contains(@class, 'el-input el-input--small el-input--suffix')]")))
        Select_type[len(Select_type)-1].click()
        # Select 'Drop point'
        Region = WebDriverWait(driver,60).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class, 'el-select-dropdown el-popper')]//span[text()='Drop Point']")))
        Region[len(Region)-1].click()
        
        # Click Inquiry button
        Inq=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
        #Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
        Inq[0].click()
        # Click Export button
        Exp=driver.find_elements(By.XPATH,"//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
        Exp[0].click()
        
        # Click Export pop up box 'Yes'
        Expp=driver.find_element(By.XPATH,"//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
        # Logging time when Export button is clicked
        Expp_T=datetime.now()
        # Click 'Download' after exporting
        Expp.click() # Return the last index no. of found button and select
        # Click on Download Center and open up Download Center pop up
        time.sleep(0.2)
        DownC=driver.find_elements(By.XPATH,"//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
        DownC[0].click()
        EFD = driver.find_element(By.XPATH,"//*[@aria-label='Export File Download']")
        
        # Needed Time and Module Name info for finding file to download in DC pop up.
        ## Date 1 (original)
        Expp_T1=Expp_T.strftime('%Y-%m-%d %H:%M:%S')
        
        ## Date 2 (with one second margin)
        Expp_T2=Expp_T+timedelta(0,1)
        Expp_T2=Expp_T2.strftime('%Y-%m-%d %H:%M:%S')
        
        # Module name
        module_name = '到件监控Summary'
        
        # Expp_T time log as conditon for row to look for
        # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")
        
        # XPath for file with Completed status
        ## Search File XPATH with original time
        SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        ## XPATH with alternative time (1 second more)
        SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        
        # Loop for 'Completed' status; repeat if no 'Completed' status
        
        while len(EFD.find_elements(By.XPATH, SFXPATH1)) == 0 and len(EFD.find_elements(By.XPATH, SFXPATH2)) == 0: # When both time as not found
                # IB = inquiry button syntax
                EFD_IB=EFD.find_elements(By.XPATH,"//button[contains(., 'Inquiry')]")
                EFD_IB[1].click()

        # Condition between two path: one with original time, the other with one second margin or error.
        ## SFXPATH_Test for match case condition
        SFXPATH_Test = len(EFD.find_elements(By.XPATH, SFXPATH1))

        match SFXPATH_Test:
            # For case of without time delay - syntax with original time    
            case SFXPATH_Test if SFXPATH_Test > 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH1) # original time syntax
            # For case of one second late - syntax with one second margin added    
            case SFXPATH_Test if SFXPATH_Test == 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH2) 
            
        SFRD = SFR[0] # Search File Row to Download

        # DB = Download Button
        DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
        DB.click() # Click Download Button
        
        Close=EFD.find_element(By.XPATH, "//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
        Close.click()
        
        # Wait file to complete download using while loop
        ## Download Folder
        
        ## Hold until Chrome download temporary file no longer present
        dl_wait = True # Default True
        while dl_wait:
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(dp_am):
                if fname.endswith('.crdownload'):
                    dl_wait = True

    # Close Delivery Monitoring > Download Center
    #DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    #DMDC_Close[len(DMDC_Close)-1].click()
    
    time.sleep(0.2)
    
    # Close Delivery Monitoring tab
    AM_Close=driver.find_elements(By.XPATH,"//div[@id='tab-ArriveMonitor|crisbiIndex' and contains(text(), 'Arrival Monitoring (New)')]//span[contains(@class, 'el-icon-close')]")
    AM_Close[0].click()

#def Parcel_Tracking():

def Departure_backup(driver,Date):  # Departure monitoring (new) -> open -> Next Station button -> date(s) # Final Version V1 complete

    # 1. Change download location for this execution to "Delivery Monitoring Folder"
    # Reference: https://stackoverflow.com/questions/22734399/python-selenium-webdriver-dynamically-change-download-directory

    driver.command_executor._commands['send_command'] = (
        'POST', '/session/$sessionId/chromium/send_command')
    ## Set up path for changing download location for this function; (optional) to be used when renaming as well.
    dp_dpm = os.path.join(os.getcwd(), "Departure Monitoring")  # dp_dm = download path - delivery monitoring
    params = {
        'cmd': 'Page.setDownloadBehavior',
        'params': {'behavior': 'allow', 'downloadPath': dp_dpm}
    }
    driver.execute("send_command", params)

    # 2. Open website
    driver.get('https://jms.jtexpress.my/indexSub')
    driver.implicitly_wait(2)
    # 2.1 Search box type in

    JMSSearch = driver.find_element(By.XPATH, "//input[@placeholder='Please check']")
    JMSSearch.send_keys("Departure Monitoring (New)")
    time.sleep(1)
    JMSSearch.send_keys(Keys.DOWN)
    JMSSearch.send_keys(Keys.ENTER)

    # Tick 'Previous Station'
    try:
        driver.find_element(By.XPATH,
                            "//div[@class='el-form-item el-form-item--small']//span[@class='el-checkbox__input' and following-sibling::span[@class='el-checkbox__label' and text()='Next Station']]").click()
    except NoSuchElementException:
        pass

    # 3.2 For loop of date, inquiry/export, and then download
    ## Preparing time box index (time box will generate a new index every time the pop up is closed)
    tbi1 = 0  # For start time
    tbi2 = 1  # For end time

    for d in Date:
        d = str(d)
        # Click 'Start time' date box
        ## assign stb and etb to same element; element has the same syntax
        stb = etb = driver.find_elements(By.XPATH,
                                         "//div[@class='el-date-editor el-input el-input--small el-input--prefix el-input--suffix el-date-editor--datetime']//input[@type='text' and @placeholder='Choose date and time']")
        stb[0].click()

        # Insert first date
        timebox = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        # print(tbi1)  print tbi1 for debugging purpose; remove '#' toe enable

        timebox[tbi1].send_keys(Keys.CONTROL + 'a')
        timebox[tbi1].send_keys(Keys.DELETE)
        timebox[tbi1].send_keys(d)
        # timebox[tbi1].send_keys(Keys.ENTER)
        # timebox[tbi1].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        if tbi1 < 1:
            tbi1 += 1
        else:
            tbi1

        # Click 'End Time' date box
        etb[1].click()
        # Insert second date
        timebox = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Select date']")))
        time.sleep(0.1)
        # tbi2=1 # tbi = time box index
        timebox[tbi2].send_keys(Keys.CONTROL + 'a')
        timebox[tbi2].send_keys(Keys.DELETE)
        timebox[tbi2].send_keys(d)
        # timebox[tbi2].send_keys(Keys.ENTER)
        # timebox[tbi2].send_keys(Keys.ENTER)
        # If to decide whether to add start time box index (note only first time will tbi1 needed to be added by one)
        # if tbi2 < 1:
        #    tbi2+=1
        # else:
        #    tbi2

        # Select type: Drop point
        ## Click drop down box
        Select_type = driver.find_elements(By.XPATH,
                                           "//span[@title='Type :']/ancestor::div[@class='el-form-item el-form-item--small']//div[contains(@class, 'el-input el-input--small el-input--suffix')]")
        Select_type[len(Select_type) - 1].click()
        ## Select 'Drop point'
        Region = driver.find_elements(By.XPATH,
                                      "//div[contains(@class, 'el-select-dropdown el-popper')]//span[text()='Drop Point']")
        Region[len(Region) - 1].click()

        # Click Inquiry button
        Inq = driver.find_elements(By.XPATH,
                                   "//button[contains(@class,'el-button btn-query el-button--primary el-button--small') and span[text()='Inquiry']]")
        # Inq[(len(Inq))-1].click() # Return the last index no. of found button and select
        Inq[0].click()
        # Click Export button
        Exp = driver.find_elements(By.XPATH,
                                   "//button[contains(@class,'el-button el-button--info el-button--small') and span[text()='Export']]")
        Exp[0].click()

        # Click Export pop up box 'Yes'
        Expp = driver.find_element(By.XPATH,
                                   "//button[contains(@class,'el-button el-button--default el-button--small el-button--primary comfirm-btn')]")
        # Logging time when Export button is clicked
        Expp_T = datetime.now()
        # Click 'Download' after exporting
        Expp.click()  # Return the last index no. of found button and select
        # Click on Download Center and open up Download Center pop up
        time.sleep(0.2)
        DownC = driver.find_elements(By.XPATH,
                                     "//button[contains(@class, 'el-button el-button--info el-button--small')]//span/span[text()='Download Center']")
        DownC[0].click()
        EFD = driver.find_element(By.XPATH, "//*[@aria-label='Export File Download']")

        # Needed Time and Module Name info for finding file to download in DC pop up.
        ## Date 1 (original)
        Expp_T1 = Expp_T.strftime('%Y-%m-%d %H:%M:%S')

        ## Date 2 (with one second margin)
        Expp_T2 = Expp_T + timedelta(0, 1)
        Expp_T2 = Expp_T2.strftime('%Y-%m-%d %H:%M:%S')

        # Module name
        module_name = '发件监控Summary'
        # print(Expp_T1)

        # Expp_T time log as conditon for row to look for
        # Rows = EFD[1].find_elements(By.XPATH, f"//tr[.//span[contains(text(), '{Expp_T}')] and .//span[contains(text(), '{module_name}')]]")

        # XPath for file with Completed status
        ## Search File XPATH with original time
        SFXPATH1 = f"//tr[.//td[contains(., '{Expp_T1}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"
        ## XPATH with alternative time (1 second more)
        SFXPATH2 = f"//tr[.//td[contains(., '{Expp_T2}')] and .//td[contains(., '{module_name}')] and .//td[contains(., 'Completed')]]"

        # Loop for 'Completed' status; repeat if no 'Completed' status

        while len(EFD.find_elements(By.XPATH, SFXPATH1)) == 0 and len(
                EFD.find_elements(By.XPATH, SFXPATH2)) == 0:  # When both time as not found
            # IB = inquiry button syntax
            EFD_IB = EFD.find_elements(By.XPATH, "//button[contains(., 'Inquiry')]")
            EFD_IB[1].click()

        # Condition between two path: one with original time, the other with one second margin or error.
        ## SFXPATH_Test for match case condition
        SFXPATH_Test = len(EFD.find_elements(By.XPATH, SFXPATH1))

        match SFXPATH_Test:
            # For case of without time delay - syntax with original time
            case SFXPATH_Test if SFXPATH_Test > 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH1)  # original time syntax
            # For case of one second late - syntax with one second margin added
            case SFXPATH_Test if SFXPATH_Test == 0:
                SFR = EFD.find_elements(By.XPATH, SFXPATH2)

        SFRD = SFR[0]  # Search File Row to Download

        # DB = Download Button
        DB = WebDriverWait(SFRD, 10).until(EC.presence_of_element_located((By.XPATH, ".//button[@title='Download']")))
        DB.click()  # Click Download Button

        Close = EFD.find_element(By.XPATH,"//div[contains(@class, 'el-dialog__header') and .//span[text()='Export File Download']]//button[@aria-label='Close']")
        Close.click()

        # Wait file to complete download using while loop
        ## Download Folder
        # DownLoc_Def=os.path.join(os.getcwd(),"Departure Monitoring")

        ## Hold until Chrome download temporary file no longer present
        dl_wait = True  # Default True
        while dl_wait:
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(dp_dpm):
                if fname.endswith('.crdownload'):
                    dl_wait = True

    # Close Delivery Monitoring > Download Center
    # DMDC_Close=driver.find_elements(By.XPATH,"//div[@class='el-dialog__header' and .//span[@class='el-dialog__title' and text()='Export File Download']]//button[@aria-label='Close']")
    # DMDC_Close[len(DMDC_Close)-1].click()

    time.sleep(0.2)

    # Close Delivery Monitoring tab
    DPM_Close = driver.find_elements(By.XPATH,
                                     "//div[@id='tab-SendOutMonitor|crisbiIndex' and contains(text(), 'Departure Monitoring (New)')]//span[contains(@class, 'el-icon-close')]")
    DPM_Close[0].click()

