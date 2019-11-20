import cx_Oracle
import sys
import xlrd
import xlwt
from xlutils import copy
import os
import fnmatch



def createConn(usr,pswd,hostname,port,sid):
	newString = usr+"/"+pswd+"@"+hostname+":"+port+"/"+sid
	print newString
	return newString


def modifyColumn(file):
	print "modify"
	print copy
	file_locn = file  #pass argument while running the script
	file_locn1= file_locn[:-1]  #convert file to .xls with additional column rank
	rb = xlrd.open_workbook(file_locn)
	read = rb.sheet_by_index(0)
	wb = copy.copy(rb)
	s = wb.get_sheet(0)
	
	## This column is left as null as it is populated by the script later 
	s.write(0,32,'Report_date')  
	
	#column number to be changed while running the script ----- IMPORTANT ----------------
	for r in range(read.nrows):
		if r == 0:
			s.write(r,33,'Rank')
		else:
			s.write(r,33,r)
			#s.write(r,32,date)
		
		
	
	wb.save(file_locn1)

#Need to read the Excel again for import


def insertIntoDB(file,conn):	
	print "insert"
	cursor = conn.cursor()
	s = xlrd.open_workbook(file)
	read = s.sheet_by_index(0)
	
	data=[[read.cell_value(r,c) for c in range(read.ncols)] for r in range(read.nrows)]
		
	insert_sql ="INSERT INTO PRJ_CLOUD_SLOW_CLICKS values ("
	
	for r in range(read.nrows-1):
		print (r+1)
		for c in range(read.ncols):
			if type(data[r+1][c]) == unicode:
				insert_sql=insert_sql+"\'"+str(data[r+1][c]).replace("'", r"''")+"\'"+","
			elif type(data[r+1][c]) == float:
				insert_sql=insert_sql+str(data[r+1][c])+","
			else:
				insert_sql=insert_sql+"\'\'"+","
		insert_sql=insert_sql[:-1]
		insert_sql+=")"
		print insert_sql + "\n"
		
		cursor.execute(insert_sql)		
		insert_sql ="INSERT INTO PRJ_CLOUD_SLOW_CLICKS values ("
		
	cursor.close()
	conn.commit()
	
	##Updating the new report bugs with reported date as otherwise association script will fail.
	#sql = "update prj_cloud_slow_clicks set report_date = (select h.rptdate from rpthead@BUGDB_PROD h where prj_cloud_slow_clicks.bug_id = h.rptno ) where exists(select 1 from rpthead@BUGDB_PROD h where prj_cloud_slow_clicks.bug_id = h.rptno) and (bug_id !='Add a bug' or bug_id is not null)"
	#cursor = conn.cursor()
	#cursor.execute(sql)
	#cursor.close()
	#conn.commit()
	
	

def loopRunner():
	#reading all access log files
	for file in os.listdir(file_dir):
		if fnmatch.fnmatch(file, '*.xlsx'):
			print "Processing file....", file
			print file_dir + file
			str = file_dir + "\\" + file 
			mainRunner(str)
			

def mainRunner(file_locn):		
	#Assigning the parameters passed while running the script	
	#file_locn = sys.argv[1]
	#from_date = sys.argv[1] 
	#to_date = sys.argv[2]
	
	modifyColumn(file_locn)
	conn =cx_Oracle.connect(createConn('user','pwd','host','port','sid'))
	conn.begin()
	
	#print file_locn
	print "Connection Established"
	insertIntoDB(file_locn[:-1],conn)
	conn.close()
	
	def getConn():
		return conn

	

#File location format printed
print "Please specify file loc"
#Date format printed
print "Please specify date in this format : 2017-02-05    2017-02-19"
print "Please execute mainRunner with the params"

file_dir = sys.argv[1]
#from_date = sys.argv[2] 
#to_date = sys.argv[2]

print file_dir
#print from_date
#print to_date

#mainRunner()
loopRunner()


	
