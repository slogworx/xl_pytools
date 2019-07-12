import datetime
import random
import pyodbc
import csv
import sys

server = 'APP-SAM-MSSQL\APP'
database = 'EOS'  # NPS

cnxn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
cursor = cnxn.cursor()

def check_file(filename):
	# The file needs to end in .qry.
	if filename[-4:] == '.qry':
		return True
	else:
		return False
		

def query_file(filename):
		
	try:
		with open(filename, 'r') as qry_in:
			query = qry_in.read()
				
	except Exception as e:
		raise e
		
	return query
	

"""
query_pull()

Submit a query to the database and save the response as a csv.
"""
def query_pull(query):
	
	try:
		cursor.execute(query)
	
	except Exception as e:
		raise e
	
	header = [column[0] for column in cursor.description]
	survey = []
	for row in cursor.fetchall():
		survey.append(dict(zip(header, row)))
	# TODO: Automatically convert csv to xlsx w/ XlsxWriter
	now = datetime.datetime.now()
	# Generate filename based on date, time, and a random number between 1000 and 2000
	timestamp = f'NPS_QUERY_{now.day}{now.month}{now.year}{now.hour}{now.minute}{now.second}_{random.randint(1000, 2000)}'
	try:
		with open(f'{timestamp}.csv', 'a', newline='') as csvout:  # TODO: generate filename from datetime
			writer = csv.DictWriter(csvout, dialect='excel', fieldnames=header)
			writer.writeheader()
			for line in survey:
				writer.writerow(line)
	except Exception as e:
		raise e


def main():

	filename = sys.argv[1]  # query file specified from cli
	
	if check_file(filename):
		query = query_file(filename)
		query_pull(query)
		print(f'{filename} has been executed and the response written to csv.')
	else:
		print(f'{filename} is not a valid query file')

	
if __name__ == "__main__":
	main()