from pathlib import Path
from typing import List
import win32com.client
import pywintypes


def get_project_root() -> str:
    return str(Path(__file__).parent.parent)


def enter_prompt(prompt: str = ''):
    """ Prompts the user with the given prompt and waits to proceed until the user hits 'Enter'

    Args:
        prompt (str): What prompt to give the user. Default ''
    """
    input(prompt)


def yes_no_prompt(prompt: str = '', yes_list: List[str] = ['y'], no_list: List[str] = ['n'],
                  strip_input: bool = False) -> bool:
    """ Prompts the user for a response returns true or false depending on user input. If input does not match anything
    in yes_list or no_list, prompt the user again.

    The user input is stripped on both ends to remove any white space. To indicate pressing 'Enter', use '' in the
    yes_list or no_list input.

    Args:
        prompt (str): What question to prompt the user for a response for. Default ''
        yes_list (List[str]): A list of input strings for which to return a True value. Default ['y']
        no_list (List[str]): A list of input strings for which to return a False value. Default ['n']
        strip_input (bool): Whether to strip the input. Default False

    Returns:
        bool: True if the user responded yes, False if no
    """

    answer = input(prompt)
    if strip_input:
        answer = answer.strip()

    while answer not in yes_list + no_list:
        answer = input(prompt).strip()
        if strip_input:
            answer = answer.strip()

    if answer in yes_list:
        return True
    else:
        return False


def close_excel_wb(excel_wb_filepath: str, save_changes: bool | None = None):
    """ If the given file is open in Excel, this function opens the file in Excel and then closes the file.

    Args:
        excel_wb_filepath (str): The abbreviated filename of this Excel workbook.
            e.g. use 'export.xlsx', not 'C:\\export.xlsx'
        save_changes (bool): Optional. Whether to save the Excel file as we close it. Default to None so that Excel
            will prompt the user to manually decide to close the file or not.
    """

    try:
        excel_wb = win32com.client.GetObject(excel_wb_filepath)
        excel_wb.Close(save_changes)
    except pywintypes.com_error:
        # If here, the file could not be found, so there is no need to close it.
        print(f'Specified file {excel_wb_filepath} could not be found. ')
        pass


close_excel_wb('A:\\zvcq11Exxport.xlsx')
