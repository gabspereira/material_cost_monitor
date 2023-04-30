import win32com.client
import sys
import subprocess
import time
from datetime import date
import pandas as pd
from utils.include_validation import Purchase
from utils.cost_block_grouper import Grouper
from utils.auxiliar_procedures import kill_excel
from utils.export_to_csv import run as export_to_csv

start_time = time.time()
pd.set_option('mode.chained_assignment', None)



# - Basic Parameters:
begin_date  = '01.01.2022'
end_date    = date.today().strftime("%d.%m.%Y")
print_date  = date.today().strftime("%Y-%m-%d")
plant       = 'BX91'
my_sap_path = "C:/temp/data/SAP/"

# - List of SuperBOM HALBS
my_list =  ['A7B82004355427', 'A7B82004467553', 'A7B82004467563']


# SAP Session:

def saplogin():

    try:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(10)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection("LP1 - Spiridon LAM - Production System", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return


        # Get all lists that compose a NXAIR SuperBOM
        for x in range(len(my_list)):

            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "/ncs12"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = my_list[x]
            session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = "BX91"
            session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = "best"
            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").caretPosition = 14
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[43]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = my_sap_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = my_list[x] + ".xlsx"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            print('List ' + my_list[x] + ' downloaded.')
            kill_excel()


        # Append all downlaoded lists in a single one:
        print('Append files start.')
        df0 = pd.read_excel(my_sap_path + my_list[0] + '.xlsx')
        df1 = pd.read_excel(my_sap_path + my_list[1] + '.xlsx')
        df2 = pd.read_excel(my_sap_path + my_list[2] + '.xlsx')
        data = [df0, df1, df2]
        print('Files listed..')
        data_appended = pd.concat(data)
        print('Files appened...')

        # Build a single list of material to start searching in SAP
        df_single = data_appended[['Component number']].drop_duplicates()
        df_single.to_clipboard(index=False)
        print('Single list of materials copied to clipboard.')

        # Get Purchase Orders using copied values from Single Lists
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nme80fn"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtP_QCOUNT").text = ""
        session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtSP$00011-LOW").text = "bx91"
        session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = begin_date
        session.findById("wnd[0]/usr/ctxtSP$00001-HIGH").text = end_date
        session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtSP$00001-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtSP$00001-HIGH").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = my_sap_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PURCHASE_ORDERS.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('File PURCHASE ORDERS was updated.')
        kill_excel()


        # Get Moving Avera Prices using copied values from Single Lists

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sie/sla_mm_cleasing"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[0]/usr/radP_CONT").select()
        session.findById("wnd[0]/usr/ctxtP_WERKS").text = "BX91"
        session.findById("wnd[0]/usr/ctxtS_MTART-LOW").text = "ABF"
        session.findById("wnd[0]/usr/ctxtS_MTART-HIGH").text = "ZVER"
        session.findById("wnd[0]/usr/radP_CONT").setFocus()
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = my_sap_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CLEASING.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('File CLEASING was updated.')
        kill_excel()


        # Get TABLE MARC using copied values from Single Lists

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sie/sla_mm_marc"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "bx91"
        session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").text = "6006"
        session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = my_sap_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "TABLE_MARC.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('File TABLE_MARC was updated.')
        kill_excel()


        # Get INFO RECORDS using copied values from Single Lists

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nme1m"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtI_EKORG-LOW").text = "609b"
        session.findById("wnd[0]/usr/ctxtI_WERKS-LOW").text = "bx91"
        session.findById("wnd[0]/usr/ctxtI_WERKS-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtI_WERKS-LOW").caretPosition = 4
        session.findById("wnd[0]/usr/btn%_IF_LIFNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()
        session.findById("wnd[0]/usr/btn%_IF_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = my_sap_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "INFO_REC.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('File INFO_REC was updated.')
        kill_excel()



        # Include into the SuperBOM: Procurement type, Bulk, Include (0/1)
        # Call the function `include_validation.py` to enrich the bill of material table

        df = Grouper(data_appended)
        df = Purchase(df)
        df.to_excel(my_sap_path + print_date + '-nxtoos-valid.xlsx', index=True)
        print('SuperBOM file saved.')



    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None

# saplogin()

    export_to_csv

    print("--- %s seconds ---" % (time.time() - start_time))
