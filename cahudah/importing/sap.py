import sys
import traceback
from typing import Literal
from cahudah.utils import get_project_root
import win32com.client
import pandas


def export_zvc09(sales_order: int, sales_line: int,
                 output_folder: str = get_project_root() + '\\data\\helper\\',
                 output_filename:str = 'zvc09_export.xlsx',
                 characteristic_type: Literal['description', 'technical', 'none'] = 'none',
                 parent_material: str = 'CUSTOM_AHU',
                 plant_num: int = 1170) -> str:
    """ Runs and exports to the specified folder and filename as a .xlsx file the output of SAP transaction ZVC09 for
    a specific job.

    ZVC09 is the SAP transaction which returns the Multi-Level BOM (Bill of Materials) Explosion.

    Args:
        sales_order (int): The given job's sales order number
        sales_line (int): The line number of the given sales order number
        output_folder (str): The string representation of the output folder. Defaults to 'cahudah/data/helper'
        output_filename (str): The string representation of the output file.
            Must end with '.xlsx'. Defaults to 'zvc09_export.xlsx'.
        characteristic_type (str): One of 'description', 'technical', or 'none'. Specifies how characteristics will be
            shown in the BOM. Default 'none'.
        parent_material (str): The SAP material number that acts as the parent material for the rest of the BOM.
            Defaults to 'CUSTOM_AHU'
        plant_num (int): The Greenheck Group plant number for the given job. Defaults to Innovent's plant number: '1170'

    Returns:
        str: Absolute filepath of the exported xlsx file

    Raises:
        AssertionError: If the user was not able to connect to SAP GUI
    """
    print(sales_order)
    print(sales_line)
    print(output_folder)
    print(output_filename)
    print(characteristic_type)
    print(parent_material)
    print(plant_num)

    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        assert isinstance(sap_gui, win32com.client.CDispatch)

        sap_app = sap_gui.GetScriptingEngine
        assert isinstance(sap_gui, win32com.client.CDispatch)

        sap_connection = sap_app.Children(0)
        assert isinstance(sap_gui, win32com.client.CDispatch)

        sap_session = sap_connection.Children(0)
        assert isinstance(sap_gui, win32com.client.CDispatch)
    except AssertionError:
        print("\nCould not connect to SAP GUI. Perhaps SAP was not open?\n")
        raise

    sap_app = sap_gui.GetScriptingEngine


    # assert isinstance(sap_app.Children, object)
    sap_connections = sap_app.Children


    x = 4

    # try:
    #     sap_gui_auto = win32com.client.GetObject("SAPGUI")
    #
    #
    #
    # except:
    #     print(sys.exc_info()[0])
    #
    # finally:
    #     pass


export_zvc09(4564564,20)