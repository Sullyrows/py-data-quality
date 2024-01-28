import pandas as pd 
from pathlib import Path
from openpyxl.workbook import Workbook
import json 


class Excel_DQ(object): 

    # read methods into class
    from .methods import (
        _read_excel,
        _row_completeness,
        _col_completeness
    )

    def __init__(
        self,
        workbook_file: Workbook | Path,
        worksheet: str | list[str] = None
    ): 
        """__init__ _summary_
        """
        # initialized values
        self.workbook_file = workbook_file
        self.worksheet = worksheet


        match self.workbook_file: 
            case Path(): 
                # when pathlib path read in 
                self._read_excel(file_path = self.workbook_file)
            case Workbook(): 
                # set the same 
                self.workbook = self.workbook_file

        # get sheets (whether it's a list or not )

        # format self.worksheet
        if worksheet is None: 
            self._sheet = self.workbook.active
        else: 
            self._sheet = self.workbook.worksheets[self.worksheet]

        self._values = self._sheet.values

        # check and validate values 
        self.row_completeness = self._row_completeness()
        self.col_completeness = self._col_completeness()


    def to_dict(self, save_location: Path = None) -> dict: 
        """to_dict dictionary generator / serializer to JSON

        Args:
            save_location (Path): If populated with Path, will save JSON dictionary. If none, just return dictionary. Defaults to None. 

        Returns:
            dict: Output dictionary
        """

        output_dict = {}
        output_dict["workbook_file"] = self.workbook_file.relative_to(Path.cwd())

        output_dict["number_rows"] = len(list(self._values))