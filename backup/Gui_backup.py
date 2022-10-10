import sys
import threading
from tkinter import *
from mainFile import run


def main_func():
    # Execute Tkinter
    root.mainloop()


# create root window
root = Tk()

# root window title and dimension
root.title("Online Barcode Scanner")

image = PhotoImage(file=r"efl3pl.png")
image = image.subsample(3, 3)

Label(root, image=image).grid(row=0, column=0,
                              columnspan=1, rowspan=1, padx=5, pady=5)

# status label settings
status = Label(root, text="Waiting for Gate Pass....")
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
                 command=lambda: run(email_var.get(), status))

btn_stop = Button(root, text="Stop",
                  command=lambda: sys.exit())
btn_run.grid(row=3, column=0, pady=2)
btn_stop.grid(row=3, column=1, pady=2)

if __name__ == '__main__':
    main_func()
