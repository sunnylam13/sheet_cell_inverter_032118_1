# Scratch Notes

## Wednesday, March 21, 2018 9:42 AM

Program actions

nested `for` loops to read data into a list of lists data structure

maybe `sheetData[x][y]` for cell at column x and row y

then writing a new sheet

use `sheetData[y][x]` for cell at column x and row y

test data

	file:///Users/sunnyair/Dropbox/python_projects/sheet_cell_inverter_032118_1/sheet_cell_inverter/updatedProduceSales.xlsx

	sheet_cell_inverter/updatedProduceSales.xlsx


	sheet_cell_inverter/updatedProduceSales_v2.xlsx

should be in same folder as main program

execution code

	python3 sheet_cell_inverter.py pathToSpreadsheet

	python3 sheet_cell_inverter.py updatedProduceSales.xlsx

	python3 sheet_cell_inverter.py updatedProduceSales_v2.xlsx

## Wednesday, March 21, 2018 10:23 AM

It appears there's a limit to the number of rows that can be converted to columns...  

When it's around 18279 rows or after that, conversion to columns results in an error...

Meaning there's a limit to the number of rows that can be inverted...

	2018-03-21 10:24:27,244 - DEBUG - The invert_cell_coordinate is:  ZZZ3
	 2018-03-21 10:24:27,244 - DEBUG - The invert_column_letter is:  ZZZ
	 2018-03-21 10:24:27,244 - DEBUG - The invert_cell_coordinate is:  ZZZ4
	Traceback (most recent call last):
	  File "/usr/local/lib/python3.6/site-packages/openpyxl/utils/cell.py", line 111, in get_column_letter
	    return _STRING_COL_CACHE[idx]
	KeyError: 18279

	During handling of the above exception, another exception occurred:

	Traceback (most recent call last):
	  File "sheet_cell_inverter.py", line 95, in <module>
	    row_analyzer(inverted_dict,1,upper_row_max + 1)
	  File "sheet_cell_inverter.py", line 78, in row_analyzer
	    invert_column_letter = get_column_letter(rowValue) # the row value becomes the column value so get its letter
	  File "/usr/local/lib/python3.6/site-packages/openpyxl/utils/cell.py", line 113, in get_column_letter
	    raise ValueError("Invalid column index {0}".format(idx))
	ValueError: Invalid column index 18279
	MacBook-Air:sheet_cell_inverter sunnyair$

In fact `updatedProduceSales.xlsx` has 23758 rows so that means 18279 really is the hard cap as far as I can tell...

That's the limit...

When I use a smaller sheet the program works fine...

