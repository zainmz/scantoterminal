import time

import xlwings as xw

from valueToTerminal import *


def run(gate_pass, status):
    url = "http://hjprod.itx360.com/webtrmgw/webtrmgw.dll?MfcISAPICommand=ProcessTerminalForm&terminal_name=WTLP01&server_name=PROD-HJ-APPS-BK&echo_type=Y&inputdata=F1"
    excel_file_loc = r'C:\Users\lka-logisticspark\Desktop\Body Line\Auto Print Lables-BODYLINE.xlsm'

    #
    # Get the open macro file instance
    #
    wb = xw.Book(excel_file_loc)

    getExcelValue(wb)
    # Delete session file
    #
    # os.remove('selenium_session')

    # start the browser with terminal URL
    browser = build_driver()
    browser.get(url)

    # Log in to the Virtual Terminal
    status.configure(text="Logging In..")
    logon(browser)
    time.sleep(10)
    # Dock in
    status.configure(text="Docking In..")
    # dockIN(browser, gate_pass, status)
    # goHome(browser)

    # Go to ASN receipt
    status.configure(text="Receiving in Progress!")
    ASNReceipt(browser, wb, gate_pass, status)
    # F3(browser)  # confirm receipt
    # F4(browser)  # re-confirm receipt
    # goHome(browser)

    # Go to Dock Out
    status.configure(text="Docking Out..")
    # dockOut(browser, gate_pass, status)

    # Logout of Virtual Terminal
    status.configure(text="Logged Out")
    logout(browser)
########################################################################################################################
from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os

from selenium.webdriver.chrome.service import Service as ChromeService  # Similar thing for firefox also!
from subprocess import CREATE_NO_WINDOW  # This flag will only be available in windows

########################################################################################################################
#
#                                               SELENIUM FUNCTIONS
#
########################################################################################################################
from getFromExcel import *

SELENIUM_SESSION_FILE = 'selenium_session'
SELENIUM_PORT = 9515

username = 'LPCHAMIKA'
password = '09876'
equipment_zone = 'LPCHAMIKA'


def build_driver():
    options = Options()
    options.add_argument("--disable-infobars")
    options.add_argument("--enable-file-cookies")
    options.add_argument("--incognito")
    options.add_argument("--disk-cache-size=0")
    options.add_experimental_option("detach", True)

    if os.path.isfile(SELENIUM_SESSION_FILE):
        session_file = open(SELENIUM_SESSION_FILE)
        session_info = session_file.readlines()
        session_file.close()

        executor_url = session_info[0].strip()
        session_id = session_info[1].strip()

        capabilities = options.to_capabilities()
        driver = webdriver.Remote(command_executor=executor_url, desired_capabilities=capabilities)
        # prevent annoying empty chrome windows
        driver.close()
        driver.quit()

        # attach to existing session
        driver.session_id = session_id
        return driver

    chrome_service = ChromeService('chromedriver')
    chrome_service.creationflags = CREATE_NO_WINDOW

    # hide the command prompt window
    driver = webdriver.Chrome(service=chrome_service, options=options, port=SELENIUM_PORT)
    driver.delete_all_cookies()

    session_file = open(SELENIUM_SESSION_FILE, 'w')
    session_file.writelines([
        driver.command_executor._url,
        "\n",
        driver.session_id,
        "\n",
    ])
    session_file.close()

    return driver


########################################################################################################################
#
#                                            VIRTUAL TERMINAL STANDARD FUNCTIONS
#
########################################################################################################################

def submit(browser):
    #
    # Submit button press
    #
    browser.find_element(By.XPATH, '//*[@value="Submit"]').click()


def F1(browser):
    #
    # Go back one page
    #
    browser.find_element(By.XPATH, '//*[@value="F1"]').click()


def F3(browser):
    #
    # Complete Inbound
    #
    browser.find_element(By.XPATH, '//*[@value="F3"]').click()


def F8(browser):
    #
    # Next page/ Skip
    #
    browser.find_element(By.XPATH, '//*[@value="F8"]').click()


def F4(browser):
    #
    # Complete Inbound Verification
    #
    browser.find_element(By.XPATH, '//*[@value="F4"]').click()


def logon(browser):
    #
    # Log in to the virtual scanner
    #
    try:

        # Check whether at the terminal details page
        check = browser.find_elements(By.NAME, 'server_name')

        # if at terminal details page, check whether details already entered otherwise enter them.
        if len(check) > 0:
            value = browser.find_element(By.NAME, 'server_name').get_attribute('value')
            if len(value) > 0:
                submit(browser)
            else:
                browser.find_element(By.NAME, "server_name").send_keys('PROD-HJ-APPS-BK')
                browser.find_element(By.NAME, "port_number").send_keys('4500')
                browser.find_element(By.NAME, "terminal_name").send_keys('WTLP01')  # WTLP01
                browser.find_element(By.XPATH, "//input[@name='save_info']").click()
                submit(browser)

        # check if not at login page, go back to login page.
        while len(browser.find_elements(By.XPATH, "//*[text()='USER ID             ']")) == 0:
            print('going back1')
            F1(browser)

        # check whether at the login page, enter the login details.
        check = browser.find_elements(By.XPATH, "//*[text()='USER ID             ']")
        print(check)
        if len(check) > 0:

            browser.find_element(By.NAME, "inputdata").send_keys(username)
            submit(browser)

            browser.find_element(By.NAME, "inputdata").send_keys(password)
            submit(browser)

            browser.find_element(By.NAME, "inputdata").send_keys(equipment_zone)
            submit(browser)

        else:
            print("Log on failed.")

    except:
        print("Log on failed.")


def logout(browser):
    #
    # Logout of the Virtual Terminal
    #
    # check if not at login page, go back to login page.
    try:
        while len(browser.find_elements(By.XPATH, "//*[text()='USER ID             ']")) == 0:
            print('going back1')
            F1(browser)
        print('logged out')
    except:
        print("Failed to log out!")


def goHome(browser):
    #
    # Go to the Home page of Virtual Terminal
    #
    try:
        while len(browser.find_elements(By.XPATH, "//*[text()='1 Dock              ']")) == 0:
            print('going back1')
            F1(browser)
        print('At Home Page')
    except:
        print("Failed to log out!")


########################################################################################################################
#
#                                            VALUE ENTERING FUNCTIONS
#
########################################################################################################################
def enterGatePassID(browser, value, status):
    #
    # Enter the gate pass ID
    #
    try:
        browser.find_element(By.NAME, "inputdata").send_keys(value)
        submit(browser)
    except:
        print("Enter the gate pass ID")


def sendValue(browser, value):
    #
    # Send barcodes to the Virtual Terminal
    #
    try:
        browser.find_element(By.NAME, "inputdata").send_keys(value)
        submit(browser)
    except:
        print("Failed to reach barcode entry page")


########################################################################################################################
#
#                                          UNLOADING  VIRTUAL TERMINAL FUNCTIONS
#
########################################################################################################################

def dockIN(browser, gate_pass, status):
    #
    # Go to dock in process page
    #
    try:
        browser.find_element(By.NAME, "inputdata").send_keys("1")  # go to Receipts
        submit(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("1")  # go to Unload
        submit(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("3")  # go to Dock in Time
        submit(browser)

        enterGatePassID(browser, gate_pass, status)  # enter the gate pass

        browser.find_element(By.NAME, "inputdata").send_keys("1")  # Enter Door Location
        submit(browser)
        F1(browser)  # go back
        browser.find_element(By.NAME, "inputdata").send_keys("1")  # Start Unload
        submit(browser)

        enterGatePassID(browser, gate_pass)  # enter the gate pass
    except:
        print("Failed to Dock In")


def ASNReceipt(browser, wb, gate_pass, status):
    #
    # ASN Receipts
    #
    try:
        browser.find_element(By.NAME, "inputdata").send_keys("1")  # go to Receipts
        submit(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("3")  # go to Vendor Receipt
        submit(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("7")  # ASN Receipt
        submit(browser)
        enterGatePassID(browser, gate_pass, status)
        while True:
            # get the carton count of the operation for repeated ASN
            if getCartonCount(wb) % 50 == 0:


            # get completion status of operation
            if getCompleteStatus(wb) == 'Complete':
                break

            # send the barcode value to HJ Virtual Terminal
            # save the current value
            current_value = getExcelValue(wb)
            sendValue(browser, current_value)

            # wait for new value
            while current_value == getExcelValue(wb):
                if current_value != getExcelValue(wb):
                    break
                else:
                    pass
    except:
        print("Failed ASN")


def dockOut(browser, gate_pass, status):
    #
    # Dock out
    #
    try:
        browser.find_element(By.NAME, "inputdata").send_keys("1")  # go to Receipts
        submit(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("1")  # go to unload
        submit(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("2")  # go to end unload
        submit(browser)

        enterGatePassID(browser, gate_pass, status)  # enter gate pass id

        browser.find_element(By.NAME, "inputdata").send_keys("4")  # go dock out time
        submit(browser)

    except:
        print("Failed ASN")
