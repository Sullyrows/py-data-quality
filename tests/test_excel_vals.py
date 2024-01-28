import pytest
import pandas as pd
from py_data_quality.excel import Excel_DQ
import pathlib

@pytest.fixture(scope="module")
def excel_dq(excel_path: pathlib.Path): 
    """test just building excel data quality"""
    yield Excel_DQ(
        workbook_file=excel_path
    )


class Test_excel_dq: 

    def test_json_gen(self, excel_dq: Excel_DQ): 
        """test json generation"""

        try: 
            excel_dq.to_dict()
        except Exception as ex: 
            pytest.fail(str(ex))