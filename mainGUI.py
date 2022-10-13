import sys
import threading
from tkinter import *
import customtkinter
import http.client as httplib

from valueToTerminal import *

url = "http://hjprod.itx360.com/webtrmgw/webtrmgw.dll?MfcISAPICommand=ProcessTerminalForm&terminal_name=WTLP01&server_name=PROD-HJ-APPS-BK&echo_type=Y&inputdata=F1"
excel_file_loc = r'C:\Users\96598\Documents\Desktop\Auto Print Lables-BODYLINE.xlsm'


def manual():
    sys.exit()


def complete():
    wb = xw.Book(excel_file_loc)
    setComplete(wb)


class App(customtkinter.CTk):
    keepGoing = False

    def __init__(self):
        super().__init__()

        customtkinter.set_appearance_mode("light")

        # Title and GUI Size
        self.iconbitmap("icon.ico")
        self.title("Online Barcode Scanning System")

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure((0, 1), weight=1)

        self.image = PhotoImage(file=r"efl3pl.png")
        self.image = self.image.subsample(2, 2)

        # Status Display & settings
        Label(master=self, image=self.image).grid(row=0, column=0, sticky=W)

        self.status = customtkinter.CTkLabel(master=self,
                                             text="Enter Login Details",
                                             height=100,
                                             fg_color=("white", "gray75"),
                                             corner_radius=8)

        self.status.configure(font=("Arial", 30))
        self.status.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=20, pady=20)

        # Login information collection
        self.username_label = customtkinter.CTkLabel(master=self,
                                                     text="Username",
                                                     height=10)
        self.username_label.configure(font=("Arial", 15))

        self.username_var = StringVar(self, value="")
        self.username = customtkinter.CTkEntry(master=self,
                                               placeholder_text="",
                                               textvariable=self.username_var,
                                               height=40,
                                               border_width=2,
                                               corner_radius=5)

        self.username_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=5, padx=20)
        self.username.grid(row=2, column=1, columnspan=2, sticky="nsew", pady=5, padx=20)

        self.password_label = customtkinter.CTkLabel(master=self,
                                                     text="Password",
                                                     height=10)
        self.password_label.configure(font=("Arial", 15))

        self.password_var = StringVar(self, value="")
        self.password = customtkinter.CTkEntry(master=self,
                                               placeholder_text="",
                                               textvariable=self.password_var,
                                               show='*',
                                               height=40,
                                               border_width=2,
                                               corner_radius=5)

        self.password_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=5, padx=20)
        self.password.grid(row=3, column=1, columnspan=2, sticky="nsew", pady=5, padx=20)

        self.fork_label = customtkinter.CTkLabel(master=self,
                                                 text="Fork ID",
                                                 height=10)
        self.fork_label.configure(font=("Arial", 15))

        self.fork_var = StringVar(self, value="")
        self.fork = customtkinter.CTkEntry(master=self,
                                           placeholder_text="",
                                           textvariable=self.fork_var,
                                           height=40,
                                           border_width=2,
                                           corner_radius=5)

        self.fork_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=5, padx=20)
        self.fork.grid(row=4, column=1, columnspan=2, sticky="nsew", pady=5, padx=20)

        self.terminal_label = customtkinter.CTkLabel(master=self,
                                                     text="Terminal ID",
                                                     height=10)
        self.terminal_label.configure(font=("Arial", 15))

        self.terminal_var = StringVar(self, value="WTLP01")
        self.terminal = customtkinter.CTkEntry(master=self,
                                               placeholder_text="",
                                               textvariable=self.terminal_var,
                                               height=40,
                                               border_width=2,
                                               corner_radius=5)

        self.terminal_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=5, padx=20)
        self.terminal.grid(row=5, column=1, columnspan=2, sticky="nsew", pady=5, padx=20)

        # BUTTONS #####################################################################

        self.start = customtkinter.CTkButton(master=self, height=40, fg_color="#90EE90", hover_color="gray",
                                             text_font=("Arial", 15),
                                             text="Start",
                                             command=self.run_thread)

        self.start.grid(row=2, column=3, sticky=E, pady=5, padx=20)

        self.stop = customtkinter.CTkButton(master=self, height=40, fg_color="#FF7F7F", hover_color="gray",
                                            text_font=("Arial", 15), text="Stop",
                                            command=self.stop_thread)

        self.stop.grid(row=3, column=3, sticky=E, pady=5, padx=20)

        self.manual = customtkinter.CTkButton(master=self, height=40, fg_color="#ADD8E6", hover_color="gray",
                                              text_font=("Arial", 15), text="Manual",
                                              command=self.destroy)

        self.manual.grid(row=4, column=3, sticky=E, pady=5, padx=20)

        self.complete = customtkinter.CTkButton(master=self, height=40, fg_color="#FCF55F", hover_color="gray",
                                                text_font=("Arial", 15), text="Complete",
                                                command=complete)

        self.complete.grid(row=5, column=3, sticky=E, pady=5, padx=20)

        # Credits settings
        self.creator = customtkinter.CTkLabel(self, text="Developed by Zain Zameer")
        self.creator.configure(font=("Courier", 10))
        self.creator.grid(row=6, column=1, sticky=W, pady=5, padx=20)

    def stop_thread(self):
        print('stopping thread')
        self.keepGoing = False

    def run_thread(self):
        self.keepGoing = True
        print('starting thread')

        # NO INPUT HANDLING
        if len(self.username_var.get()) == 0:
            self.status.configure(text='Enter Username')
            return
        if len(self.password_var.get()) == 0:
            self.status.configure(text='Enter Password')
            return
        if len(self.fork_var.get()) == 0:
            self.status.configure(text='Enter Fork ID')
            return
        if len(self.terminal_var.get()) == 0:
            self.status.configure(text='Enter Terminal ID')
            return

        threading.Thread(target=run, args=(self.username_var.get(), self.password_var.get(),
                                           self.fork_var.get(), self.terminal_var.get(),
                                           self.status, self.keepGoing)).start()


########################################################################################################################
#
#                                            MAIN RUNNING CODE
#
########################################################################################################################

def is_stop_pressed(browser, status, keepGoing):
    status.configure(text="Stopping...")
    if not keepGoing:
        goHome(browser)
        logout(browser)
        status.configure(text="Stopped!")
        sys.exit()


def checkInternetHttplib(url="www.google.com", timeout=3):
    connection = httplib.HTTPConnection(url, timeout=timeout)
    try:
        # only header requested for fast operation
        connection.request("HEAD", "/")
        connection.close()  # connection closed
        print("Internet On")
        return True
    except Exception as exep:
        print(exep)
        return False


def run(username, password, fork, terminal, status, keepGoing):
    try:
        print('thread started')

        if not checkInternetHttplib("www.google.com", 3):
            status.configure(text="No Internet")
            return

        #
        # Get the open macro file instance
        #
        wb = xw.Book(excel_file_loc)

        if getGatePassID(wb) == "None":
            status.configure(text="No GatePass ID")
            return

        if getLocation(wb) == "None":
            status.configure(text="No Unloading Location Set")
            return

        if getDoorLocation(wb) == "None":
            status.configure(text="No Door Location Set")
            return

        getExcelValue(wb)
        # Delete session file
        #
        # os.remove('selenium_session')

        # start the browser with terminal URL
        browser = build_driver()
        browser.get(url)

        # Log in to the Virtual Terminal
        status.configure(text="Logging In..")
        logon(browser, username, password, fork, terminal, status)

        is_stop_pressed(browser, status, keepGoing)
        # Dock in
        status.configure(text="Docking In..")
        dockIN(browser, status, wb)
        status.configure(text="Dock In Complete!.")
        is_stop_pressed(browser, status, keepGoing)
        goHome(browser)

        is_stop_pressed(browser, status, keepGoing)
        # Go to ASN receipt
        status.configure(text="Receiving in Progress!")
        ASNReceipt(browser, wb, status)
        status.configure(text="Receiving Complete!")
        is_stop_pressed(browser, status, keepGoing)
        goHome(browser)

        is_stop_pressed(browser, status, keepGoing)
        # Go to Dock Out
        status.configure(text="Docking Out..")
        dockOut(browser, getGatePassID(wb), status)
        status.configure(text="Dock out Complete!")

        # Logout of Virtual Terminal
        status.configure(text="Logged Out")
        logout(browser)

    except Exception as e:
        print(e)
        status.configure(text=e)


if __name__ == "__main__":
    app = App()
    app.mainloop()
