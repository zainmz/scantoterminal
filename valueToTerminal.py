from bs4 import BeautifulSoup
from selenium import webdriver
import win32com.client as wincom

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
import os

from selenium.webdriver.chrome.service import Service as ChromeService  # Similar thing for firefox also!
from subprocess import CREATE_NO_WINDOW  # This flag will only be available in windows

########################################################################################################################
#
#                                               SELENIUM FUNCTIONS
#
########################################################################################################################


from getFromExcel import *

SELENIUM_SESSION_FILE = './selenium_session'
SELENIUM_PORT = 9515

#username = 'LPCHAMIKA'
#password = '09876'
#equipment_zone = 'LPCHAMIKA'

page_source = ''


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

    driver = webdriver.Chrome(options=options, port=SELENIUM_PORT)
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


def getPage(browser):
    #
    # get the current page from selenium
    #
    global page_source
    page_source = browser.page_source


def display(browser, status):
    #
    # Get the text displayed on the terminal
    #
    getPage(browser)
    soup = BeautifulSoup(page_source, "html.parser")
    text = soup.get_text("|", strip=True)
    text_list = text.split('|')

    print(text_list)
    status.configure(text=text_list[2])


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


def logon(browser, username, password, fork, terminal, status):
    #
    # Log in to the virtual scanner
    #
    try:

        # ERROR HANDLING
        # Check if server down
        # check if not at login page, go back to login page.
        while len(browser.find_elements(By.XPATH, "//*[text()='Advantage Workflow Engine is down.']")) != 0:
            status.configure(text='Server Down!')
            browser.refresh()

        # check if terminal already in use
        while len(browser.find_elements(By.XPATH, "//*[text()='Terminal is already in use.']")) != 0:
            status.configure(text='Terminal Unavailable!')
            browser.refresh()

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
                browser.find_element(By.NAME, "terminal_name").send_keys(terminal)  # WTLP01
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

            browser.find_element(By.NAME, "inputdata").send_keys(fork)
            submit(browser)

        # check if terminal already in use
        while len(browser.find_elements(By.XPATH, "//*[text()='INVALID USER ID     ']")) != 0:
            status.configure(text='Invalid Username!')
            logout(browser)
            return True

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
        while len(browser.find_elements(By.XPATH, "//*[text()='1 Receipts          ']")) == 0:
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

def dockIN(browser, status, wb):
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

        enterGatePassID(browser, getGatePassID(wb), status)  # enter the gate pass
        sendValue(browser, getDoorLocation(wb))  # send the door location
        submit(browser)  # confirm transaction complete

    except:
        print("Failed to Dock In")


def ASNReceipt(browser, wb, status):
    #
    # ASN Receipts
    #
    try:
        while getCompleteStatus(wb) != 'Complete':

            speak = wincom.Dispatch("SAPI.SpVoice")

            # go to ASN Receipt
            browser.find_element(By.NAME, "inputdata").send_keys("1")  # go to Receipts
            submit(browser)
            browser.find_element(By.NAME, "inputdata").send_keys("3")  # go to Vendor Receipt
            submit(browser)
            browser.find_element(By.NAME, "inputdata").send_keys("5")  # ASN Receipt
            submit(browser)
            enterGatePassID(browser, getGatePassID(wb), status)  # Enter the gate pass ID

            while len(browser.find_elements(By.XPATH, "//*[text()='LP Scanned : 50     ']")) == 0:

                if getCompleteStatus(wb) == 'Complete':
                    break

                text = "Start Scanning"
                speak.Speak(text)
                status.configure(text='Start Scanning!')
                status.config(fg='Green')
                # send the barcode value to HJ Virtual Terminal
                current_value = getExcelValue(wb)
                sendValue(browser, getExcelValue(wb))
                submit(browser)  # confirm quantity

                try:
                    display(browser, status)  # show the LP numbers scanned on gui
                except:
                    pass

                if len(browser.find_elements(By.XPATH, "//*[text()='LP Scanned : 50     ']")) != 0:
                    break

                # wait for new value
                while current_value == getExcelValue(wb):
                    if getCompleteStatus(wb) == 'Complete':
                        break
                    if current_value != getExcelValue(wb):
                        break
                    else:
                        pass

                # if len(browser.find_elements(By.XPATH, "//*[text()='LP Scanned : 10     ']")) != 0:
                # break
            text = "Stop Scanning"
            speak.Speak(text)
            status.configure(text='Stop Scanning!')
            status.config(fg='Red')
            F3(browser)

            while len(browser.find_elements(By.XPATH, "//*[text()='Advantage Workflow Engine is down.']")) != 0:
                print('reloading page')
                browser.refresh()

            F4(browser)
            sendValue(browser, getLocation(wb))
            submit(browser)

            goHome(browser)

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
        submit(browser)

        F1(browser)
        browser.find_element(By.NAME, "inputdata").send_keys("4")  # go dock out time
        submit(browser)
        enterGatePassID(browser, gate_pass, status)


    except:
        print("Failed ASN")
