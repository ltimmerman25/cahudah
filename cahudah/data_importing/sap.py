import sys
import traceback
from typing import Literal, List
from cahudah.utils import get_project_root, yes_no_prompt, close_excel_wb
import win32com.client
import pywintypes
import pandas
import os.path


def initiate_sapgui_session() -> win32com.client.CDispatch:
    """ Initiates a scripting connection to an already running SAP GUI window.

    Note that SAP GUI must be opened on the user's computer for this program to work. If SAP GUI is not open, this
    function will prompt the user to open it.

    Returns:
        win32com.client.CDispatch: The COM dispatch for this SAP sap_session.
    """

    try:
        sap_gui = win32com.client.GetObject("SAPGUI")  # SAP is not open at all
        sap_app = sap_gui.GetScriptingEngine  # SAP is not open at all
        sap_connection = sap_app.Children(0)  # We have not selected any connections (eg PRD)
        sap_session = sap_connection.Children(0)
    except pywintypes.com_error:
        print('Could not connect to SAP GUI.')
        print('Ensure that you have opened and logged into a SAP GUI.')
        if yes_no_prompt('To try again, hit enter. To quit, enter q: ', [''], ['q']):
            return initiate_sapgui_session()
        else:
            sys.exit()
    return sap_session



def export_zvc09(sales_order: int, sales_line: int,
                 output_folder: str = get_project_root() + '\\data\\helper\\',
                 output_filename: str = 'zvc09_export.xlsx',
                 char_type: Literal['description', 'technical', 'none'] = 'none',
                 parent_material: str = 'CUSTOM-AHU',
                 plant_num: int = 1170) -> str:
    """ Runs and exports to the specified folder and filename as a .xlsx file the output of SAP transaction ZVC09 for
    a specific job.

    ZVC09 is the SAP transaction which returns the Multi-Level BOM (Bill of Materials) Explosion.

    Args:
        sales_order (int): The given job's sales order number
        sales_line (int): The line number of the given sales order number
        output_folder (str): The string representation of the output folder. Defaults to 'cahudah\\data\\helper'
        output_filename (str): The string representation of the output file.
            Must end with '.xlsx'. Defaults to 'zvc09_export.xlsx'.
        char_type (str): One of 'description', 'technical', or 'none'. Specifies how characteristics will be
            shown in the BOM. Default 'none'.
        parent_material (str): The SAP material number that acts as the parent material for the rest of the BOM.
            Defaults to 'CUSTOM_AHU'
        plant_num (int): The Greenheck Group plant number for the given job. Defaults to Innovent's plant number: '1170'

    Returns:
        str: Absolute filepath of the exported xlsx file

    Raises:
        TypeError: For improper function parameters
        pywintypes.com_error: When SAP is for whatever reason unable to complete the specified transaction.
    """

    sap_session = initiate_sapgui_session()

    try:
        # Initiate zvc09 transaction
        sap_session.findById('wnd[0]/tbar[0]/okcd').text = '/nzvc09'
        sap_session.findById('wnd[0]').sendVKey(0)

        # Set BOM input parameters
        match char_type:
            case 'none':
                sap_session.findById('wnd[0]/usr/radP_NONE').select()
            case 'description':
                sap_session.findById('wnd[0]/usr/radP_CHAR').select()
            case 'technical':
                sap_session.findById('wnd[0]/usr/radP_TECH').select()
            case _:
                raise TypeError('char_type does not match one of the specified options.')
        sap_session.findById('wnd[0]/usr/ctxtA_MATNR').text = parent_material
        sap_session.findById('wnd[0]/usr/ctxtA_PLANT').text = plant_num
        sap_session.findById('wnd[0]/usr/txtA_VBELN').text = sales_order
        sap_session.findById('wnd[0]/usr/txtA_POSNR').text = sales_line
        sap_session.findById('wnd[0]/tbar[1]/btn[8]').press()

        # Export BOM to Excel file
        sap_session.findById('wnd[0]/usr/cntlG_CONTAINER/shellcont/shell').pressToolbarContextButton('&MB_EXPORT')
        sap_session.findById('wnd[0]/usr/cntlG_CONTAINER/shellcont/shell').selectContextMenuItem('&XXL')
        sap_session.findById('wnd[1]/usr/radRB_OTHERS').select()
        sap_session.findById('wnd[1]/tbar[0]/btn[0]').press()

        # Enter filepath, name, and export it
        sap_session.findById('wnd[1]/usr/ctxtDY_PATH').text = output_folder
        sap_session.findById('wnd[1]/usr/ctxtDY_FILENAME').text = output_filename
        if os.path.isfile(output_folder + output_filename):
            # If this file already exists, first close it if open, then click the button to overwrite the old file
            close_excel_wb(output_folder + output_filename)
            sap_session.findById('wnd[1]/tbar[0]/btn[11]').press()
        else:
            # If this file does not exist yet, create a new file
            sap_session.findById('wnd[1]/tbar[0]/btn[0]').press()

    except Exception:
        print('\nSAP GUI was accessed, but was not able to run zvc09.')
        print('Inspect the error code for more information.\n')
        raise


export_zvc09(9076969,20)


