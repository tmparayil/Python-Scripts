import cx_Oracle
import sys
import xlrd
import xlwt
from xlutils import copy

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
	
	

def invokingPackage(fro,to,conn): 
	
	print "invoke"
	#ConnSudheer DB connection established
	#Currently not handling DB connection issues
	cursor = conn.cursor()
	
	##Not Tested
	cursor.callproc("dbms_output.enable")
	
	
	
	#getting from_date and to_date from python script execution arguments
	from_date = fro 
	to_date = to
	
	cursor.callproc('CLOUD_SLOW_CLICK.SET_FROM_DATE',[from_date])  #setting from_date in the package
	cursor.callproc('CLOUD_SLOW_CLICK.SET_TO_DATE',[to_date])		#setting to_date in the package
	
	statusVar = cursor.var(cx_Oracle.NUMBER)
	lineVar = cursor.var(cx_Oracle.STRING)
	while True:
		cursor.callproc("dbms_output.get_line", (lineVar, statusVar))
		if statusVar.getvalue() != 0:
			break
		print lineVar.getvalue()
	
	cursor.callproc('CLOUD_SLOW_CLICK.REMOVE_DUPLICATE_ADD_BUGS')
	cursor.execute('commit')
	
	cursor.callproc('CLOUD_SLOW_CLICK.REMOVE_DUPLICATES')
	cursor.execute('commit')
	
	statusVar = cursor.var(cx_Oracle.NUMBER)
	lineVar = cursor.var(cx_Oracle.STRING)
	while True:
		cursor.callproc("dbms_output.get_line", (lineVar, statusVar))
		if statusVar.getvalue() != 0:
			break
		print lineVar.getvalue()
	
	
	cursor.callproc('CLOUD_SLOW_CLICK.LOG_MAIN')
	cursor.execute('commit')
    
	statusVar = cursor.var(cx_Oracle.NUMBER)
	lineVar = cursor.var(cx_Oracle.STRING)
	while True:
		cursor.callproc("dbms_output.get_line", (lineVar, statusVar))
		if statusVar.getvalue() != 0:
			break
		print lineVar.getvalue()
		
	cursor.close()
	

def executeAssociation():	
	print "execute"
#	sql = "select family,product,view_id,region_view_id,click_id,component_type,action_type,display_text,bug_id from prj_cloud_slow_clicks where to_date(report_date,'DD-MM-YYYY')=to_date(sysdate,'DD-MM-YYYY') OR (bug_id is not null AND report_date is null)"
#	
#	result = cursor.execute(sql)
#	
#	#loop through result set and write into excel
#	#for data in result:
		

def mainRunner():		
	#Assigning the parameters passed while running the script	
	#file_locn = sys.argv[1]
	#from_date = sys.argv[1] 
	#to_date = sys.argv[2]
	
	modifyColumn(file_locn)
	conn =cx_Oracle.connect(createConn('fusion','fusion','slc02hme.us.oracle.com','1522','slc02hme'))
	conn.begin()
	
	
	print "Connection Established"
	insertIntoDB(file_locn[:-1],conn)
	invokingPackage(from_date,to_date,conn)
	executeAssociation()
	conn.close()
	
	def getConn():
		return conn

	

#File location format printed
print "Please specify file locn: 		Ex: C:\\Users\\tparayil.ORADEV\\Downloads\\TopCloudSlowClicksReport_CLOUD_ProjectsDomain_PRJ_ALL_Production_2017-02-20.xlsx"
#Date format printed
print "Please specify date in this format : 2017-02-05    2017-02-19"
print "Please execute mainRunner with the params"

file_locn = sys.argv[1]
from_date = sys.argv[2] 
to_date = sys.argv[3]

print file_locn
print from_date
print to_date

mainRunner()


	