import sys
import threading
from tkinter import *
from test import *
import pyttsx3

keepGoing = False


def main_func():
    # Execute Tkinter
    root.mainloop()


def run_thread():
    global keepGoing
    keepGoing = True
    print('starting thread')
    threading.Thread(target=run, args=(email_var.get(), status)).start()


def stop_thread():
    print('stopping thread')
    global keepGoing
    keepGoing = False


# create root window
root = Tk()

# root window title and dimension
root.title("Online Barcode Scanner")

root.geometry("500x200")

image = PhotoImage(file=r"efl3pl.png")
image = image.subsample(3, 3)

Label(root, image=image).grid(row=0, column=0,
                              columnspan=1, rowspan=1, padx=5, pady=5)

# status label settings
status = Label(root, text="Waiting for Gate Pass....")
status.config(font=("Arial", 15))
status.grid(row=1, column=1, sticky=W, pady=5)

# Credits settings
creator = Label(root, text="Developed by Zain Zameer")
creator.config(font=("Courier", 6))
creator.grid(row=7, column=0, sticky=W, pady=5)

# all widgets will be here
email_to = Label(root, text="GATE PASS ID")

# entry widgets
email_to.grid(row=2, column=0, sticky=W, pady=5)

# declaring string variable
# for storing name and password
email_var = StringVar(root, value="")

# entry widgets, used to take entry from user
e1 = Entry(root, textvariable=email_var)

# this will arrange entry widgets
e1.grid(row=2, column=1, pady=2)

btn_run = Button(root, text="Run",
                 command=run_thread)

btn_stop = Button(root, text="Stop",
                  command=stop_thread)

btn_run.grid(row=3, column=0, pady=2)
btn_stop.grid(row=3, column=1, pady=2)


########################################################################################################################
#
#                                            MAIN RUNNING CODE
#
########################################################################################################################
def is_stop_pressed(browser, status):
    global keepGoing
    status.configure(text="Stopping...")
    if not keepGoing:
        goHome(browser)
        logout(browser)
        status.configure(text="Stopped!")
        sys.exit()


def run(gate_pass, status):
    try:
        print('thread started')
        url = "http://hjprod.itx360.com/webtrmgw/webtrmgw.dll?MfcISAPICommand=ProcessTerminalForm&terminal_name=WTLP01&server_name=PROD-HJ-APPS-BK&echo_type=Y&inputdata=F1"
        excel_file_loc = r'C:\Users\lka-logisticspark\Desktop\Body Line\Auto Print Lables-BODYLINE.xlsm'

        # start the text to voice engine
        engine = pyttsx3.init()

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

        is_stop_pressed(browser, status)
        # Dock in
        status.configure(text="Docking In..")
        dockIN(browser, gate_pass, status, wb)
        status.configure(text="Dock In Complete!.")
        goHome(browser)

        is_stop_pressed(browser, status)
        # Go to ASN receipt
        status.configure(text="Receiving in Progress!")
        ASNReceipt(browser, wb, gate_pass, status, engine)
        status.configure(text="Receiving Complete!")
        goHome(browser)

        is_stop_pressed(browser, status)
        # Go to Dock Out
        status.configure(text="Docking Out..")
        # dockOut(browser, gate_pass, status)
        status.configure(text="Dock out Complete!")

        # Logout of Virtual Terminal
        status.configure(text="Logged Out")
        logout(browser)

    except Exception as e:
        status.configure(text=e)


if __name__ == '__main__':
    main_func()
