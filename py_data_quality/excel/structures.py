from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd 
import numpy as np 
from dataclasses import dataclass, field 


@dataclass
class ExcelTable(object): 
    """class for handling table reading"""

    cell_range: str = field(init=True)
    col_range: list[int] = field(init=True)
    row_range: list[int] = field(init=True)
    col_row: int = field(init=True)



class DiscoverTable(object): 

    from .methods import _row_completeness

    def __init__(self, worksheet: Worksheet): 
        """initialize class"""
        
        # intialize values 
        self.worksheet = worksheet 
        self._values = self.worksheet.values
        # initialize defaults
        self.tables = None

        # build inference structure for class
        _ws_tables = self.worksheet.tables
        if _ws_tables is not None: 
            pass