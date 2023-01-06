import sys

import mysql.connector
from tkinter import *
from pandas.io import sql as sql
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os
import pandas as pd
import tkdnd
from pyepc import SGTIN
from pyepc.exceptions import DecodingError


class Store:
    def __init__(self, store_num, date_input, cycle, cycle_output, item_file, matching, total_items, expected, combined):
        self.store_num = store_num
        self.date_input = date_input
        self.cycle = cycle
        self.cycle_output = cycle_output
        self.item_file = item_file
        self.matching = matching
        self.total_items = total_items
        self.expected = expected
        self.combined = combined

    def set_cycle(self, cycle_path):
        self.cycle = cycle_path

    def set_cycle_output(self, cycle_output_path):
        self.cycle_output = cycle_output_path

    def set_item_file(self, item_file_path):
        self.item_file = item_file_path

    def set_matching(self, matching_df):
        self.matching = matching_df

    def set_total_items(self, total_items_df):
        self.total_items = total_items_df

    def set_expected(self, expected_df):
        self.expected = expected_df

    def set_combined(self, combined):
        self.combined = combined

    def set_store_num(self, store_number):
        self.store_num = store_number

    def set_date_input(self, inputted_date):
        self.date_input = inputted_date

    def get_cycle(self):
        return self.cycle

    def get_cycle_output(self):
        return self.cycle_output

    def get_item_file(self):
        return self.item_file

    def get_matching(self):
        return self.matching

    def get_total_items(self):
        return self.total_items

    def get_expected(self):
        return self.expected

    def get_combined(self):
        return self.combined

    def get_store_num(self):
        return self.store_num

    def get_date_input(self):
        return self.date_input

    def toString(self):
        string = "Store Number: " + str(self.store_num) \
        + "\n\tDate: " + str(self.date_input) \
        + "\n\tCycle Count path: " + str(self.cycle) \
        + "\n\tCycle Count Output path: " + str(self.cycle_output) \
        + "\n\tItem File paths: " + str(self.item_file[0]) + "\n\t\t\t\t" + str(self.item_file[1]) \
        + "\n\tMatching Data Frame: " + str(self.matching) \
        + "\n\tTotal Items Data Frame: " + str(self.total_items) \
        + "\n\tExpected Items Data Frame: " + str(self.expected) \
        + "\n\tCombined Items Data Frame: " + str(self.combined)
        return string