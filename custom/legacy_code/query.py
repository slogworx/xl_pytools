""" Query specified SQL database from service and output query to a formatted Excel (xlsx) workbook.

"""

import xlsxwriter
import datetime
import random
import pyodbc
import sys

from write_xlsx import write_xlsx

server = 'APP-SAM-MSSQL\\APP'
database = 'EOS'  # Transactional NPS

cnxn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
cursor = cnxn.cursor()
		

def query_file(filename):  # Reads query file
		
	try:
		with open(filename, 'r') as qry_in:
			query = qry_in.read()
	except Exception as e:
		print(f'{e}.')
		
	return query
	

def query_pull(query='', output_file='default_output.xlsx'):
	
	workbook = xlsxwriter.Workbook(f'{output_file}.xlsx')
	worksheet = workbook.add_worksheet(output_file)

	try:
		cursor.execute(query)
	
	except Exception as e:
		raise e
	
	# Create the header
	col_names = []
	cell_loc = { 'row': 0, 'col': 0 }
	for col_name in cursor.description:
		col_names.append(col_name[0])
	
	# Write the header
	write_xlsx(col_names, cell_loc, workbook, worksheet, header = True)
	
	cell_loc['row'] = 1
	cell_loc['col'] = 0
	for row in cursor.fetchall():
		write_xlsx(row, cell_loc, workbook, worksheet, header = False)
		cell_loc['row'] += 1
	
	workbook.close()


def main():
	
	filename = sys.argv[1]  # query file specified from cli
	query = query_file(filename)
	if query:
		query_pull(query, filename[:-4])
		print(f'{filename} has been executed and the response written to {filename[:-4]}.output.xlsx.')
	else:
		print(f'Please use a valid query file.')

	
if __name__ == "__main__":
	if len(sys.argv) < 2:
		print(f'Usage: python {sys.argv[0]} [path/filename.qry]')
	else:
		main()