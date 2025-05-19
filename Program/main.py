from csv import excel
from operator import truediv
from Part1 import Part1 as part_1
import os
import sys



from openpyxl import load_workbook
from pathlib import Path

curr_dir = Path(__file__).parent
main_dir = curr_dir.parent
excel_dir = main_dir / "Excel"
excel_estimate_path = excel_dir / "_kalkulacja edit.xlsx"


### Part 0 ###
def check_if_excel_open(path):
    try:
        with open(path, "r+"):
            return False
    except IOError:
        return True

if check_if_excel_open(excel_estimate_path):
    print("Excel opened, close it to start program")
    sys.exit()

### Part 1 ###
elem_id = input("Element ID: ")
part_1.main(elem_id)


os.startfile(excel_estimate_path)

















