import sys

import mysql.connector
from tkinter import *
from pandas.io import sql as sql
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Progressbar
from tkinter import filedialog
import os
import pandas as pd
# import tkdnd
import tkinterDnD
from pyepc import SGTIN
from pyepc.exceptions import DecodingError
import DataCollection
import Store


class Interface_Creation:
    def __init__(self, root, w, h):
        self.root = root
        self.width = w
        self.height = h
        self.store_list = []
        self.store_num = None
        self.date_input = None
        self.store = Store.Store(None, None, None, None, None, None, None, None, None, None, None)
        self.folder_created = False


    def reset(self):
        self.store_list.append(self.store)
        self.store = Store.Store(None, None, None, None, None, None, None, None, None, None, None)
        entry.destroy()
        label.destroy()
        label1.destroy()
        store_num.destroy()
        date_input.destroy()
        button.destroy()
        button1.destroy()
        button3.destroy()
        quit_button.destroy()
        new_store_button.destroy()
        self.creation()

    def creation(self):
        self.root.title("In Store Supplier Compliance Report Generator")
        w = self.width
        h = self.height
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw / 2) - (w / 2)
        y = (sh / 2) - (h / 2)
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.root.config(bg="#CCE1F2")


        var1 = tk.IntVar()
        var2 = tk.IntVar()
        var3 = tk.IntVar()

        # Create an Entry widget to accept User Input
        global entry
        entry = Entry(self.root, width=40)
        entry.focus_set()
        entry.place(relx=.5, rely=.35, anchor=CENTER)

        """
        input_store_num() Function:
            The input_store_num function takes in a store number from the user, displays it on the interface, and stores it in
                a global variable to be accessed throughout the program.
            Once the user submits the store number, the input field is cleared.
        """
        def input_store_num():
            string = entry.get()
            global store_num
            store_num = Label(self.root, text=string, font=("Arial", 20, "bold"))
            store_num.config(bg="#CCE1F2", fg="#000000")
            var2.set(1)
            global store_num_input
            store_num_input = store_num['text']
            store_num.configure(text="Store: " + store_num_input)
            self.store_num = store_num_input
            self.store.store_num = store_num_input
            store_num.place(relx=.3, rely=.1, anchor=CENTER)
            label.destroy()
            entry.delete(0, END)

        """
        input_date() Function:
            The input_date function takes in a date in the format "MMDDYYYY" from the user, displays it on the interface, 
                and stores it in a global variable to be accessed throughout the program.
            Once the user submits the date, the input field is destroyed and a button to select Cycle Count files is created.
        """
        def input_date():
            string = entry.get()
            global date_input
            date_input = Label(self.root, text=string, font=("Arial", 20, "bold"))
            date_input.config(bg="#CCE1F2", fg="#000000")
            var1.set(1)
            global user_date_input
            user_date_input = date_input['text']
            date_input.configure(text="Date: " + user_date_input)
            self.date_input = user_date_input
            self.store.date_input = user_date_input
            date_input.place(relx=.7, rely=.1, anchor=CENTER)

            if not self.folder_created:
                folder_path = os.path.join(os.path.expanduser("~"), "Desktop", "TrackingReports_{0}".format(user_date_input))
                os.mkdir(folder_path)
                self.folder_created = True

            entry.destroy()
            label1.destroy()

        def close():
            self.store_list.append(self.store)
            generate_combined(self.store_list)
            self.root.destroy()

        global quit_button
        quit_button = ttk.Button(self.root, text="Quit", width=8, command=close)
        quit_button.place(relx=.5, rely=.9, anchor=CENTER)

        global label
        label = Label(self.root, text="Enter a store number (100, 355, 5625)", font=("Arial", 20, "bold"))
        label.place(relx=.5, rely=.2, anchor=CENTER)
        label.config(bg="#CCE1F2", fg="#000000")

        # Create a Button to validate Entry Widget
        global button
        button = ttk.Button(self.root, text="Submit Store Number", width=20, command=input_store_num)
        button.place(relx=.5, rely=.5, anchor=CENTER)

        # Waits for user to submit a store number
        print("Waiting for Store Number Input...")
        button.wait_variable(var2)
        print("Store Number Input Received.")
        button.destroy()

        # Create a Button to submit the desired and inputted date.
        global button3
        button3 = ttk.Button(self.root, text="Submit Date", width=10, command=input_date)
        button3.place(relx=.5, rely=.5, anchor=CENTER)

        # Display a label which prompts the user to enter a date.
        global label1
        label1 = Label(self.root, text="Enter a date in the format MMDDYYYY", font=("Arial", 20, "bold"))
        label1.place(relx=.5, rely=.2, anchor=CENTER)
        label1.config(bg="#CCE1F2", fg="#000000")
        # user_date_input = label1['text']

        # Waits for user to submit a date
        print("Waiting for Date Input...")
        button3.wait_variable(var1)
        print("Date Input Received.")
        button3.destroy()

        data = DataCollection.Data_Collection(store_num_input, user_date_input, self.root, self)
        # Create a Button to enter the select_txt() function and select the Cycle Count .txt files
        def select_txt():
            button1.destroy()
            cycle_output = data.select_txt()
            var3.set(1)
            return cycle_output

        global button1
        button1 = ttk.Button(self.root, text="Select Cycle Count Files", command=select_txt)
        button1.place(relx=.5, rely=.5, anchor=CENTER)

        print("Waiting for Cycle Count File Selection...")
        button1.wait_variable(var3)


        def generate_combined(store_list):
            data.sql_connect_combined(store_list)

        global new_store_button
        new_store_button = ttk.Button(self.root, text="New Store", command=self.reset)
        new_store_button.place(relx=.5, rely=.5, anchor=CENTER)

        self.root.mainloop()