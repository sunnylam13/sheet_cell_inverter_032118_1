# -*- coding: utf-8 -*-

#! python3

# USAGE
# python3 sheet_cell_inverter.py pathToSpreadsheet
# python3 sheet_cell_inverter.py updatedProduceSales.xlsx

import openpyxl, sys

try:
	from openpyxl.cell import column_index_from_string,get_column_letter
except ImportError:
	from openpyxl.utils import column_index_from_string,get_column_letter

import logging
logging.basicConfig(level=logging.DEBUG, format=" %(asctime)s - %(levelname)s - %(message)s")
# logging.disable(logging.CRITICAL)

