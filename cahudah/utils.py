from pathlib import Path
from typing import List


def get_project_root() -> str:
    return str(Path(__file__).parent.parent)


def yes_no_input(prompt: str = '', yes_list: List[str] = ['y'], no_list: List[str] = ['n'], strip_input: bool = False) -> bool:
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



print(yes_no_input("hiya: ", ['y', 'yer'], ['']))