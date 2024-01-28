import pytest 
from importlib import resources 
import pathlib 

@pytest.fixture(scope="session")
def excel_path() -> pathlib.Path: 
    """generate excel path"""

    with resources.path("py_data_quality","example") as f: 
        file_path = f.parents[1] / "tests/data/excel_cases.xlsx"

        return file_path
    
