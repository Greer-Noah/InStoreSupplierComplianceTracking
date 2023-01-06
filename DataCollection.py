import sys
from tkinter.ttk import Progressbar

import mysql.connector
from tkinter import *
from pandas.io import sql as sql
import tkinter as tk
from tkinter import ttk
import xlsxwriter
from tkinter import filedialog
import os
import pandas as pd
import tkdnd
from pyepc import SGTIN
from pyepc.exceptions import DecodingError
from sqlalchemy import create_engine
import pymysql
pymysql.install_as_MySQLdb()

import Store

class Data_Collection:

    def __init__(self, store_num, date_input, root, interface_creation):
        self.store_num = store_num
        self.date_input = date_input
        self.store = Store.Store(self.store_num, self.date_input, None, None, None, None, None, None, None)
        self.root = root
        self.interface_creation = interface_creation

    """
    select_txt() Function:
        The select_txt function prompts the user to select a group of Cycle Count text files.
        The function then compiles the EPCs listed in the text files into a list, of which duplicate EPCs are removed.
        This list of EPCs is then written into a single column in a new Cycle Count Excel File (.xlsx), 
            named "Store[Store #]CC[MMDDYYYY]", which is then saved on to the user's Desktop.
        The Cycle Count file location is then passed into the DecodeCycleCount(file_location).
    """
    def select_txt(self):
        epcList = []
        # --------------Prompts User for Cycle Count files--------------------------------------------------------------
        pop_up_title = "Select Store {} Cycle Count Files".format(self.store_num)
        filenames = filedialog.askopenfilenames(initialdir = "/", title = pop_up_title,
                                               filetypes = (("txt files", "*.txt"), ("all files", "*.*")))
        self.interface_creation.store.set_item_file(filenames)
        print("Cycle Count Files Selected.")

        # --------------Reads Cycle Count Files-------------------------------------------------------------------------
        for filename in filenames:
            f = open(filename, "r")
            lines = f.readlines()
            for x in lines:
                epcList.append(x.split('\n')[0])
            f.close()
        # --------------Deletes EPC duplicates--------------------------------------------------------------------------
        epcList_noDupe = [*set(epcList)]

        # --------------Exports Cycle Count input file------------------------------------------------------------------
        df1 = pd.DataFrame(epcList_noDupe, columns=['EPCs'])
        cc_file_name = "Store{0}CC{1}.xlsx".format(self.store_num, self.date_input)
        path1 = os.path.join(os.path.expanduser("~"), "Desktop/TrackingReports_{0}".format(self.store.date_input),
                             cc_file_name)
        str(path1)
        df1.to_excel(path1, sheet_name='EPCs', startrow=0, startcol=0, index=False)
        self.interface_creation.store.set_cycle(path1)


        # --------------Starts decoding process-------------------------------------------------------------------------
        # return path1
        print("Preparing to Decode...")
        self.decodeCycleCount(path1)

    """
    DecodeCycleCount(file_location) Function:
        The DecodeCycleCount(file_location) function takes in the location of the input Cycle Count file previously created
            and reads the file into a list.
        The function then attempts to decode each EPC in the list.
        If the EPC, or SGTIN, successfully decodes into a UPC, or GTIN, then that GTIN is stored in a UPC list.
        Every EPC that could not be properly decoded into a UPC is then recorded in an Error EPC list.
        This EPC list has a corresponding Error Message which is recorded into the Error UPC list.
        The function then creates a Cycle Count Output Excel file (.xlsx) with three corresponding sheets:
            1. A list of all of the inputted EPCs and their corresponding UPCs. (Sheet Name: 'Unique EPCs, Dupe UPCs')
            2. A list of all of the non-duplicated UPCs. (Sheet Name: 'Unique EPCs, Unique UPCs')
            3. A list of all of the incorrectly decoded EPCs and their corresponding error messages. (Sheet Name: 'Errors')
        This output file is saved onto the user's Desktop.
        Lastly, the function passes the 'duplicates-included' UPC list into the sql_connect_populate(upcList) function.
    """

    def decodeCycleCount(self, file_location):

        # --------------Reads in the input Cycle Count------------------------------------------------------------------
        df2 = pd.read_excel(file_location)

        epcList1 = []

        columns = df2.columns.tolist()

        # --------------Creates EPC list using input CC file------------------------------------------------------------
        for _, i in df2.iterrows():
            for c in columns:
                epcList1.append(i[c])

        upcList = []
        errorEPCs = []
        errorUPCs = []

        # --------------Decoding Process--------------------------------------------------------------------------------
        print("Decoding...")
        for x in epcList1:
            try:
                upcList.append(SGTIN.decode(x).gtin)  # Actual decode command
            except DecodingError as e:  # Handles decoding errors
                errorEPCs.append(x)
                errorUPCs.append(e)
            except TypeError as t:  # Handles decoding errors
                errorEPCs.append(x)
                errorUPCs.append(t)

        # --------------Removes Error EPCs from EPC list----------------------------------------------------------------
        for epc in errorEPCs:
            if epc in epcList1:
                epcList1.remove(epc)

        epcList2 = []
        upcList2 = []

        # --------------Creates distinct UPC list-----------------------------------------------------------------------
        for i in range(len(upcList)):
            if upcList[i] not in upcList2:
                upcList2.append(upcList[i])
                epcList2.append(epcList1[i])

        # --------------Deletes leading 0s from UPCs--------------------------------------------------------------------
        for y in range(len(upcList)):
            upcList[y] = upcList[y].lstrip('0')

        # --------------Deletes leading 0s from UPCs--------------------------------------------------------------------
        for z in range(len(upcList2)):
            upcList2[z] = upcList2[z].lstrip('0')

        # --------------Formats and prepares lists as Pandas DataFrames to be exported to .xlsx-------------------------
        df3 = pd.DataFrame(epcList1, columns=['EPCs'])
        df4 = pd.DataFrame(upcList, columns=['UPCs'])
        # df5 = pd.DataFrame(epcList2, columns=['EPCs'])  # EPC list on 'Unique EPC, Unique UPC' sheet
        df6 = pd.DataFrame(upcList2, columns=['UPCs'])
        df7 = pd.DataFrame(errorEPCs, columns=['EPCs'])
        df8 = pd.DataFrame(errorUPCs, columns=['UPCs'])


        # --------------Exports Cycle Count Output file with Duplicate UPCs, Unique UPCs, and Errors--------------------
        cc_file_name = "Store{0}CC{1}.xlsx".format(self.store_num, self.date_input)
        path10 = os.path.join(os.path.expanduser("~"), "Desktop/TrackingReports_{0}".format(self.store.date_input),
                              cc_file_name)  # Saves on Desktop
        str(path10)
        path2 = path10.split('.')
        path2.insert(-1, '_output.')  # Adds '_output' to filename
        path3 = ''.join(path2)
        writer = pd.ExcelWriter(path3, engine='xlsxwriter')

        print("Creating Cycle Count Output File...")
        df3.to_excel(writer, sheet_name='Unique EPCs, Dupe UPCs', startrow=0, startcol=0, index=False)
        # df5.to_excel(writer, sheet_name='Unique EPCs, Unique UPCs', startrow=0, startcol=0, index=False)
        df4.to_excel(writer, sheet_name='Unique EPCs, Dupe UPCs', startrow=0, startcol=1, index=False)
        df6.to_excel(writer, sheet_name='Unique EPCs, Unique UPCs', startrow=0, startcol=0, index=False)
        df7.to_excel(writer, sheet_name='Errors', startrow=0, startcol=0, index=False)
        df8.to_excel(writer, sheet_name='Errors', startrow=0, startcol=1, index=False)

        writer.save()

        str(path3)
        self.interface_creation.store.set_cycle_output(path3)
        # print(self.store.get_cycle_output())
        print("Preparing to connect to MySQL...")
        self.sql_connect_populate(upcList, self.root)

    def sql_connect_populate(self, upcList, root):

        """
        select_item_file() Function:
            The select_item_file() function prompts the user to select a group of Item File '.csv' files (Apparel and GM)
                for a respective store.
            The function then creates a single ItemFile table within the mySQL database and compiles the user's item files,
                one after the other.
        """

        def select_item_file():

            # --------------Prompts User for Item File files----------------------------------------------------------------
            pop_up_title = "Select Store {} Item Files".format(self.store_num)
            itemfiles = filedialog.askopenfilenames(initialdir="/", title=pop_up_title,
                                                    filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
            df = pd.read_csv(itemfiles[0])
            df_headers = df.head(0).columns.tolist()

            # --------------Creates Item File table-------------------------------------------------------------------------
            statement_drop_IF = "DROP TABLE IF EXISTS ItemFile"
            cursor.execute(statement_drop_IF)
            statement_headers = "CREATE TABLE ItemFile(store_number int, REPL_GROUP_NBR int, gtin bigint, ei_onhand_qty int, " \
                                "SNAPSHOT_DATE text, UPC_NBR bigint, UPC_DESC text, ITEM1_DESC text, dept_nbr int, " \
                                "DEPT_DESC text, MDSE_SEGMENT_DESC text, MDSE_SUBGROUP_DESC text, ACCTG_DEPT_DESC text, " \
                                "DEPT_CATG_GRP_DESC text, DEPT_CATEGORY_DESC text, DEPT_SUBCATG_DESC text, VENDOR_NBR int, " \
                                "VENDOR_NAME text, BRAND_OWNER_NAME text, BRAND_FAMILY_NAME text)"
            cursor.execute(statement_headers)

            # --------------Loads both Item Files into single ItemFile table------------------------------------------------
            for itemfile in itemfiles:
                itemfile_corrected = itemfile.replace(" ", "\\ ")

                statement2 = "LOAD DATA LOCAL INFILE \'{}\' " \
                             "INTO TABLE ItemFile " \
                             "CHARACTER SET latin1 " \
                             "FIELDS TERMINATED BY \',\' " \
                             "ENCLOSED BY \'\"\' " \
                             "LINES TERMINATED BY \'\\r\\n\' " \
                             "IGNORE 1 ROWS;".format(itemfile_corrected)
                var.set(1)
                cursor.execute(statement2)

        # --------------Creates mySQL Connection------------------------------------------------------------------------
        conn = mysql.connector.connect(user='root', password='password', host='127.0.0.1', database='reportsystem',
                                       allow_local_infile=True)
        cursor = conn.cursor()
        print("Connected to MySQL...")

        # --------------Waits for Item File selection-------------------------------------------------------------------
        var = tk.IntVar()
        button2 = ttk.Button(root, text="Select Item Files", command=select_item_file)
        button2.grid()
        button2.place(relx=.5, rely=.5, anchor=CENTER)

        print("Waiting for Item File Selection...")
        button2.wait_variable(var)
        print("Item Files selected.")
        button2.destroy()

        # --------------Creates UPC Drop Table------------------------------------------------------------------------------
        print("Creating UPC Drop Table...")
        stmt = "DROP TABLE if exists UPCDrop;"
        cursor.execute(stmt)
        stmt1 = "CREATE TABLE if not exists UPCDrop(UPCs bigint);"
        cursor.execute(stmt1)

        cursor.executemany("""INSERT INTO UPCDrop(UPCs) VALUES (%s) """, list(zip(upcList)))
        conn.commit()
        print('Data entered successfully.')

        # --------------Creates Matching With Apparel Table-----------------------------------------------------------------
        stmt2 = "DROP TABLE if exists MatchingWithApparel;"
        cursor.execute(stmt2)
        print("Creating Matching with Apparel Table...")
        stmt3 = "CREATE TABLE MatchingWithApparel AS " \
                "SELECT DISTINCT * FROM ItemFile " \
                "WHERE gtin IN (SELECT UPCs FROM UPCDrop);"
        cursor.execute(stmt3)

        # --------------Creates Matching Table------------------------------------------------------------------------------
        stmt4 = "DROP TABLE if exists Matching;"
        cursor.execute(stmt4)
        print("Creating Matching Table...")

        stmt5 = "CREATE TABLE Matching AS " \
                "SELECT " \
                "MatchingWithApparel.gtin," \
                "MatchingWithApparel.DEPT_CATG_GRP_DESC," \
                "MatchingWithApparel.DEPT_CATEGORY_DESC," \
                "MatchingWithApparel.VENDOR_NBR, " \
                "MatchingWithApparel.VENDOR_NAME," \
                "MatchingWithApparel.BRAND_FAMILY_NAME," \
                "MatchingWithApparel.dept_nbr, " \
                "MatchingWithApparel.REPL_GROUP_NBR " \
                "FROM MatchingWithApparel " \
                "WHERE gtin IN (SELECT UPCs FROM UPCDrop) AND dept_nbr IN ('7','9','14','17','20','22','71','72','74','87');"
        cursor.execute(stmt5)

        # --------------Creates Total Items with Apparel Table--------------------------------------------------------------
        stmt6 = "DROP TABLE IF EXISTS TotalItems_Sub;"
        cursor.execute(stmt6)
        print("Creating Total Items with Apparel Table...")
        stmt7 = "CREATE TABLE TotalItems_Sub AS " \
                "SELECT " \
                "ItemFile.ACCTG_DEPT_DESC," \
                "ItemFile.BRAND_FAMILY_NAME," \
                "ItemFile.BRAND_OWNER_NAME," \
                "ItemFile.REPL_GROUP_NBR," \
                "ItemFile.DEPT_CATEGORY_DESC," \
                "ItemFile.DEPT_CATG_GRP_DESC," \
                "ItemFile.DEPT_DESC," \
                "ItemFile.dept_nbr," \
                "ItemFile.DEPT_SUBCATG_DESC," \
                "ItemFile.ei_onhand_qty," \
                "ItemFile.gtin," \
                "ItemFile.ITEM1_DESC," \
                "ItemFile.MDSE_SEGMENT_DESC," \
                "ItemFile.MDSE_SUBGROUP_DESC," \
                "ItemFile.SNAPSHOT_DATE," \
                "ItemFile.store_number," \
                "ItemFile.UPC_DESC," \
                "ItemFile.UPC_NBR," \
                "ItemFile.VENDOR_NAME," \
                "ItemFile.VENDOR_NBR " \
                "FROM ItemFile " \
                "INNER JOIN UPCDrop  ON UPCDrop.UPCs = ItemFile.gtin " \
                "WHERE UPCDrop.UPCs = ItemFile.gtin;"
        cursor.execute(stmt7)

        # --------------Creates Total Items (GM) table----------------------------------------------------------------------
        stmt8 = "DROP TABLE IF EXISTS TotalItems_GM;"
        cursor.execute(stmt8)
        print("Creating Total Items (GM) Table...")
        stmt9 = "CREATE TABLE TotalItems_GM AS " \
                "SELECT " \
                "TotalItems_Sub.gtin," \
                "TotalItems_Sub.DEPT_CATG_GRP_DESC," \
                "TotalItems_Sub.DEPT_CATEGORY_DESC," \
                "TotalItems_Sub.VENDOR_NBR, " \
                "TotalItems_Sub.VENDOR_NAME," \
                "TotalItems_Sub.BRAND_FAMILY_NAME," \
                "TotalItems_Sub.dept_nbr " \
                "FROM TotalItems_Sub " \
                "WHERE dept_nbr IN ('7','9','14','17','20','22','71','72','74','87');"
        cursor.execute(stmt9)

        # --------------Creates Unique Vendors table------------------------------------------------------------------------
        print("Creating Unique Vendors Table...")
        stmt10 = "DROP TABLE IF EXISTS UniqueVendors;"
        cursor.execute(stmt10)
        stmt11 = "CREATE TABLE UniqueVendors" \
                 "(D7 int, D9 int, D14 int, D17 int, D20 int, D22 int, D71 int, D72 int, D74 int, D87 int);"
        cursor.execute(stmt11)

        stmt12 = "INSERT INTO UniqueVendors (D7) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('7');"
        cursor.execute(stmt12)
        stmt13 = "INSERT INTO UniqueVendors (D9) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('9');"
        cursor.execute(stmt13)
        stmt14 = "INSERT INTO UniqueVendors (D14) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('14');"
        cursor.execute(stmt14)
        stmt15 = "INSERT INTO UniqueVendors (D17) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('17');"
        cursor.execute(stmt15)
        stmt16 = "INSERT INTO UniqueVendors (D20) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('20');"
        cursor.execute(stmt16)
        stmt17 = "INSERT INTO UniqueVendors (D22) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('22');"
        cursor.execute(stmt17)
        stmt18 = "INSERT INTO UniqueVendors (D71) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('71');"
        cursor.execute(stmt18)
        stmt19 = "INSERT INTO UniqueVendors (D72) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('72');"
        cursor.execute(stmt19)
        stmt20 = "INSERT INTO UniqueVendors (D74) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('74');"
        cursor.execute(stmt20)
        stmt21 = "INSERT INTO UniqueVendors (D87) " \
                 "SELECT DISTINCT " \
                 "Matching.VENDOR_NBR " \
                 "FROM Matching " \
                 "WHERE dept_nbr IN ('87');"
        cursor.execute(stmt21)

        # --------------Creates OH Data table-------------------------------------------------------------------------------
        print("Creating OH Data Table...")
        stmt22 = "DROP TABLE IF EXISTS OHData;"
        cursor.execute(stmt22)

        stmt24 = "CREATE TABLE OHData AS " \
                 "SELECT DISTINCT gtin, ei_onhand_qty, dept_nbr FROM ItemFile " \
                 "WHERE (ei_onhand_qty > 0) AND dept_nbr IN ('7', '9', '14', '17', '20', '22', '71', '72', '74', '87');"
        cursor.execute(stmt24)

        # --------------Creates OH Data by Department Sums table------------------------------------------------------------
        print("Creating OH Data by Department Table...")
        stmt25 = "DROP TABLE IF EXISTS OHData_Dept_Sums;"
        cursor.execute(stmt25)
        stmt26 = "CREATE TABLE OHData_Dept_Sums AS SELECT OHData.dept_nbr, " \
                 "SUM(OHData.ei_onhand_qty) AS ei_onhand_qty FROM OHData GROUP BY OHData.dept_nbr " \
                 "ORDER BY OHData.dept_nbr;"
        cursor.execute(stmt26)
        stmt27 = "INSERT INTO OHData_Dept_Sums SELECT * FROM OHData_Dept_Sums " \
                 "UNION SELECT 0 dept_nbr, SUM(ei_onhand_qty) FROM OHData_Dept_Sums;"
        cursor.execute(stmt27)

        # --------------Exports Weekly Report with Matching, Total Items, and Expected Items tables-------------------------
        match_file_name = "WeeklyReport{1}.xlsx".format(self.store_num, self.date_input)
        path4 = os.path.join(os.path.expanduser("~"), "Desktop", match_file_name)
        str(path4)

        # wb = Workbook()
        global writer
        # writer = pd.ExcelWriter(path4, engine='openpyxl')

        # print("Exporting Matching File to .xlsx...")
        print("Exporting Matching...")
        df10 = sql.read_sql('SELECT * FROM Matching', conn)
        self.interface_creation.store.set_matching(df10)
        matching_sheet_name = "Matching {}".format(self.store_num)
        str(matching_sheet_name)
        # df10.to_excel(writer, sheet_name=matching_sheet_name, startrow=0, startcol=0, index=False)

        # print("Exporting Total Items File to .xlsx...")
        print("Exporting Total Items...")
        df11 = sql.read_sql('SELECT * FROM TotalItems_GM', conn)
        self.interface_creation.store.set_total_items(df11)
        total_items_sheet_name = "Total Items {}".format(self.store_num)
        str(total_items_sheet_name)
        # df11.to_excel(writer, sheet_name=total_items_sheet_name, startrow=0, startcol=0, index=False)

        # print("Exporting OH Data File to .xlsx...")
        print("Exporting OH Data...")
        df12 = sql.read_sql('SELECT * FROM OHData_Dept_Sums', conn)
        self.interface_creation.store.set_expected(df12)
        onhand_data_sums_sheet_name = "OH Data by Dept {}".format(self.store_num)
        str(onhand_data_sums_sheet_name)
        # df12.to_excel(writer, sheet_name=onhand_data_sums_sheet_name, startrow=0, startcol=0, index=False)

        # writer.save()

        # --------------Closes mySQL Connection-----------------------------------------------------------------------------
        conn.close()
        if (conn):
            conn.close()
            print("\nThe SQLite connection is closed.")


    def sql_connect_combined(self, store_list):
        user = 'root'
        pw = ''
        host = 'localhost'
        port = 3306
        db = 'ItemFile'

        db_data = 'mysql+mysqldb://' + 'root' + ':' + 'password' + '@' + '127.0.0.1' + ':3306/' \
                  + 'ItemFile' + '?charset=latin1'
        engine = create_engine(db_data)

        connection = mysql.connector.connect(user='root', password='password', host='127.0.0.1', database='ItemFile',
                                             allow_local_infile=True)

        cursor = connection.cursor(buffered=True)

        cursor.execute("DROP TABLE IF EXISTS CombinedMatching_Dupes;")
        stmt = "CREATE TABLE CombinedMatching_Dupes LIKE Matching;"
        cursor.execute(stmt)

        for store in store_list:
            store.get_matching().to_sql('CombinedMatching_Dupes', con=engine, if_exists='append', index=False)

        cursor.execute("DROP TABLE IF EXISTS CombinedMatching;")

        stmt1 = "CREATE TABLE CombinedMatching AS SELECT DISTINCT gtin, " \
                "MAX(DEPT_CATG_GRP_DESC) AS DEPT_CATG_GRP_DESC, " \
                "MAX(DEPT_CATEGORY_DESC) AS DEPT_CATEGORY_DESC, " \
                "MAX(VENDOR_NBR) AS VENDOR_NBR, " \
                "MAX(VENDOR_NAME) AS VENDOR_NAME, " \
                "MAX(BRAND_FAMILY_NAME) AS BRAND_FAMILY_NAME, " \
                "MAX(dept_nbr)AS dept_nbr, " \
                "MAX(REPL_GROUP_NBR) AS REPL_GROUP_NBR " \
                "FROM CombinedMatching_Dupes GROUP BY gtin;"
        cursor.execute(stmt1)

        cursor.execute("DROP TABLE IF EXISTS CombinedMatching_Dupes;")

        print("Gathering Combined Matching...")
        df13 = sql.read_sql('SELECT * FROM CombinedMatching', connection)
        self.interface_creation.store.set_combined(df13)

        engine.dispose()
        connection.close()

