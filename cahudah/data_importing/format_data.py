from cahudah.utils import get_project_root


def sqlize_bom(zvc09_export_filepath: str = get_project_root() + '\\data\\helper\\zvc09_export.xlsx'):
    """ Taking the zvc09 export Excel file in cahudah\\data\\helper, formats the information to match our
    SQL implementation, removes any unnecessary information, and adds material parent info.

    Args:
        zvc09_export_filepath (str): The filepath of the zvc09 export
    """
    

    pass

