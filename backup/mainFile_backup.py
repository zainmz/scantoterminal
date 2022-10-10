import sys
import time

import xlwings as xw

from valueToTerminal import *


def run(gate_pass, status):
    try:
        url = "http://hjprod.itx360.com/webtrmgw/webtrmgw.dll?MfcISAPICommand=ProcessTerminalForm&terminal_name=WTLP01&server_name=PROD-HJ-APPS-BK&echo_type=Y&inputdata=F1"
        excel_file_loc = r'C:\Users\96598\Documents\Desktop\Auto Print Lables-BODYLINE.xlsm'

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
        time.sleep(1)
        # Dock in
        status.configure(text="Docking In..")
        dockIN(browser, gate_pass, status, wb)
        goHome(browser)

        # Go to ASN receipt
        status.configure(text="Receiving in Progress!")
        ASNReceipt(browser, wb, gate_pass, status)
        goHome(browser)

        # Go to Dock Out
        status.configure(text="Docking Out..")
        dockOut(browser, gate_pass, status)

        # Logout of Virtual Terminal
        status.configure(text="Logged Out")
        logout(browser)

    except Exception as e:
        status.configure(text=e)

