import pyodbc

from write_xlsx import write_xlsx
from query import query_pull

server = 'APP-SAM-MSSQL\\APP'
database = 'EOS'  # Transactional NPS

cnxn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
cursor = cnxn.cursor()

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
					
	respondent = input("Enter the NPS survey's 10-digit respondent ID#: ")
	if check_id(respondent):
		survey_query = f"""SELECT QST.questionNumber, QST.questionText, ANS.answerText, 
							ORD.respondentId, ORD.workorder, SRV_RSP.completeDateTimeLocal, 
							ORD.customername, SRV_CMT.comment
							FROM T_SURVEY_ANSWER SRV_ANS 
							INNER JOIN T_ANSWER ANS
							on ANS.answerId = SRV_ANS.optionId
							INNER JOIN T_SURVEY_RESPONSE SRV_RSP
							on SRV_RSP.respondentId = SRV_ANS.respondentId
							INNER JOIN T_QUESTION QST
							on QST.questionId = SRV_ANS.questionId
							INNER JOIN T_WORK_ORDER ORD
							on ORD.respondentId = SRV_ANS.respondentId
							INNER JOIN T_SURVEY_COMMENT SRV_CMT
							on SRV_CMT.respondentId = ORD.respondentId
							WHERE SRV_CMT.respondentId = {respondent}"""
		query_pull(survey_query, f'NPS_SURVEY_{respondent}')
		print(f'Transactional NPS survey #{respondent} has been saved as an Excel workbook.')
	else:
		print(f'{respondent} is an invalid respondent id. Pleaes check the number and try again.')

### TODO: Rewrite with openpyxl and combine it with mass nps tool
	
if __name__ == "__main__":
	main()
