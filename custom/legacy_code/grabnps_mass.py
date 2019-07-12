### TODO: Rewrite with openpyxl and combine it with npsgrab tool

import pyodbc
import csv
import sys

server = 'APP-SAM-MSSQL\APP'
database = 'EOS'  # NPS database

# Connect
cnxn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;unicode_results=True;CHARSET=UTF8')
cursor = cnxn.cursor()


def get_list(filename):
		
	try:
		with open(filename, 'r') as list_in:
			id_list = list_in.read()
				
	except Exception as e:
		raise e
	
	return id_list

"""
write_nps_answers()

Pull a specific NPS survey from the database and write it to csv.
"""
def write_nps_survey(respondent_id, query):
	
	if not len(respondent_id):  # If there's an empty newline in the file
		return
		
	try:
		print(f'Grabbing NPS survey# {respondent_id}')
		cursor.execute(f'{query}{respondent_id}') #  adds respondent_id argument to query
	except Exception as e:
		raise e
		
	header = [column[0] for column in cursor.description]
	survey = []
	for row in cursor.fetchall():
		survey.append(dict(zip(header, row)))
	# TODO: Automatically convert csv to xlsx w/ XlsxWriter
	try:
		with open(f'NPS_SURVEY_{respondent_id}.csv', 'a', newline='', encoding='utf8') as csvout:
			writer = csv.DictWriter(csvout, dialect='excel', fieldnames=header)
			writer.writeheader()
			for line in survey:
				writer.writerow(line)
	except Exception as e:
		if e == UnicodeEncodeError:
			pass
		else:
			raise e

"""
check_id()

Confirms the inputted respondent ID is a 10 digit number. If not, returns False. 
"""	
def check_id(n):
	if len(n) == 10:
		try:
			float(n)
			return True
		except ValueError:
			return False

def main():

	details_query = f"""SELECT ORD.respondentId, ORD.customername, SRV_RSP.completeDateTimeLocal, ORD.workorder, SRV_CMT.comment
					FROM NPS.T_WORK_ORDER ORD
					INNER JOIN NPS.T_SURVEY_RESPONSE SRV_RSP
					on SRV_RSP.workorder = ORD.workorder
					INNER JOIN NPS.T_SURVEY_COMMENT SRV_CMT
					on SRV_CMT.respondentId = ORD.respondentId
					WHERE ORD.respondentId = """
	answers_query = f"""SELECT QST.questionNumber, QST.questionText, ANS.answerText
					FROM T_SURVEY_ANSWER SRV_ANS 
					INNER JOIN T_ANSWER ANS
					on ANS.answerId = SRV_ANS.optionId
					INNER JOIN T_SURVEY_RESPONSE SRV_RSP
					on SRV_RSP.respondentId = SRV_ANS.respondentId
					INNER JOIN T_QUESTION QST
					on QST.questionId = SRV_ANS.questionId
					WHERE SRV_ANS.respondentId = """
					
	filename = sys.argv[1]
	id_list = get_list(filename).split('\n')
	
	for id in id_list:
		write_nps_survey(id, details_query)
		write_nps_survey(id, answers_query)

	print('All surveys written.')
	
	
if __name__ == "__main__":
	main()
