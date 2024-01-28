from pathlib import Path
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import pandas as pd
import numpy as np

def _read_excel(self, file_path: Path) -> Workbook: 
    """_read_excel read workbook in from Workbook

    Args:
        file_path (Path): the file path to read in

    Returns:
        Workbook: openpyxl.Workbook
    """

    # set attr for workbook for workbook 
    self.workbook = load_workbook(file_path)

    return self.workbook

# def _infer_table_values(self, )

def _col_completeness(self) -> pd.DataFrame: 
    """_col_completeness compute completeness of values in each column of df

    Returns:
        pd.DataFrame: outbound dataframe of value completeness
    """
     # cell values loaded in
    _cell_values = self._values

    # check for cell completeness as dataframe of values indexed by value array 
    value_df = pd.DataFrame(_cell_values)

    column_completeness = pd.DataFrame(
        value_df.isnull().sum()/len(value_df))\
            .reset_index()\
            .rename({0: "comleteness","index":"column"}, axis='columns')

    return column_completeness


def _row_completeness(self ) -> pd.DataFrame: 
    """_cell_completeness generated cell completeness

    Returns:
        pd.DataFrame: dataframe of completeness of values
    """
    # cell values loaded in
    _cell_values = self._values

    # check for cell completeness as dataframe of values indexed by value array 
    value_df = pd.DataFrame(_cell_values)


    row_complete = pd.DataFrame(value_df.isnull().sum(axis=1)/len(value_df))\
        .reset_index()\
        .rename({0: "completeness","index": "row"}, axis='columns')

    row_complete