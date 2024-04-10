from typing import Literal
import pandas


def export_zvc09(sales_order: int, sales_line: int, output_filepath : str = '', characteristics: Literal['description', 'technical', 'none'] = 'none',
                 parent_material: str = 'CUSTOM_AHU', plant_num: int = 1170) -> str:
    """ Runs and exports to the specified folder and filename as a .xlsx file the output of SAP transaction ZVC09 for
    a specific job.

    ZVC09 is the SAP transaction which returns the Multi-Level BOM (Bill of Materials) Explosion.

    Args:
        sales_order (int): The given job's sales order number
        sales_line (int): The line number of the given sales order number
        characteristics (str): One of 'description', 'technical', or 'none'. Specifies how characteristics will be
            shown in the BOM. Default 'none'.
        parent_material (str): The SAP material number that acts as the parent material for the rest of the BOM.
            Defaults to 'CUSTOM_AHU'
        plant_num (int): The Greenheck Group plant number for the given job. Defaults to Innovent's plant number: '1170'

    Returns:
        object:
    """
