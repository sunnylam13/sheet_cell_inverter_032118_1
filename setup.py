try:
	from setuptools import setup
except ImportError:
	from distutils.core import setup

config = {
	'description': 'The program inverts the row and column of the cells in the spreadsheet.  So a value at row 5, column 3 is moved to row 3, column 5 and vice versa.  The program affects all cells in the spreadsheet.',
	'author': 'Sunny Lam',
	'url': 'https://github.com/sunnylam13/sheet_cell_inverter_032118_1',
	'download_url': 'https://github.com/sunnylam13/sheet_cell_inverter_032118_1',
	'author_email': 'sunny.lam@gmail.com',
	'version': '0.1',
	'install_requires': ['nose'],
	'packages': ['openpyxl'],
	'scripts': [],
	'name': 'Spreadsheet Cell Inverter'
}

setup(**config)