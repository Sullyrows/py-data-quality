from openpyxl import Workbook, load_workbook
from importlib import resources 
from pathlib import Path
import json


with resources.path("py_data_quality","example") as f: 
    test_data = f.parents[1] / "tests/data"
    my_path = test_data / "excel_tests.json"

    json_values = json.loads(my_path.read_text())
