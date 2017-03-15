# in prj_cloud_slow_clicks table 
## run query to select family,product,vid,rvid,cid,ctp,atp,dt,bug_id ### where to_date(report_date)=to_date('<sysdate>','DD-MM-YYYY')
## and export to .xlsx file using sql developer


### SQL query - copy paste
## select family,product,view_id,region_view_id,click_id,component_type,action_type,display_text,bug_id from prj_cloud_slow_clicks where to_date(report_date,'DD-MM-YYYY')=to_date(sysdate,'DD-MM-YYYY') OR (bug_id is not null AND report_date is null)

### Query to update the report date from bugDB in ConnSudheer
#### update prj_cloud_slow_clicks set report_date = (select h.rptdate from rpthead@BUGDB_PROD h where prj_cloud_slow_clicks.bug_id = h.rptno ) where exists(select 1 from rpthead@BUGDB_PROD h where prj_cloud_slow_clicks.bug_id = h.rptno)

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

import xlrd

## Data Needed from table
#family
#product
#vid
#rvid
#click_id
#component
#action_type
#display_text
#bug_id

#SSO username & password
l_username = ""
l_password = ""

file_location="C:\Python27\sel scripts\cloudexport.xlsx" 
workbook=xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
data=[[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

def sso_login_check():
	return "Single Sign On" in repr(driver.title)
	
#Entire data of the query has been assigned to a 2D array.
def bug_associate_loop(i):

	r=i
	
	while(r < sheet.nrows):
		
		if r!=0:    #1st row contains row column headers
			product = ""
			
			url="https://fed.oracle.com/LogTracer/faces/UAMConfig?"+"&vid="+data[r][2]+"&rvid="+data[r][3]+"&cid="+data[r][4]+"&ctp="+data[r][5]+"&atp="+data[r][6]+"&dt="+data[r][7]+""
			driver=webdriver.Chrome();
			driver.get(url)
						
			print url+ " \t "+ str(r)
			print "loop"+ str(r)
			print data[r][8]
					
			if   data[r][1] == 'BILLING' or data[r][1] == 'PJB' :
				product="Oracle Fusion Project Billing"
			elif data[r][1] == 'COLLABORATION' or data[r][1] == 'PJL' :
				product="Oracle Fusion Project Collaboration"
			elif data[r][1] == 'CONTROL' or data[r][1] == 'PJO' :
				product="Oracle Fusion Project Control"
			elif data[r][1] == 'COSTING' or data[r][1] == 'PJC' :
				product="Oracle Fusion Project Costing"
			elif data[r][1] == 'FOUNDATION' or data[r][1] == 'PJF' :
				product="Oracle Fusion Project Foundation"
			elif data[r][1] == 'INTEGRATION' or data[r][1] == 'PJG' :
				product="Oracle Fusion Project Integration Gateway"
			elif data[r][1] == 'PROJECTMANAGEMENT' or data[r][1] == 'PJT' :
				product="Oracle Fusion Project Management"
			elif data[r][1] == 'MANAGEMENTCONTROL' or data[r][1] == 'PJE' :
				product="Oracle Fusion Project Management Control"
			elif data[r][1] == 'PERFORMANCEREPORTING' or data[r][1] == 'PJS' :
				product="Oracle Fusion Project Performance Reporting"
			elif data[r][1] == 'PORTFOLIOANALYSIS'  :
				product="Oracle Fusion Project Portfolio Analysis"
			elif data[r][1] == 'RESOURCEMANAGEMENT' or data[r][1] == 'PJR' :
				product="Oracle Fusion Project Resource Management"
			elif data[r][1] == 'AR' :
				product="Oracle Fusion Receivables"
			elif data[r][1] == 'GL' :
				product="Oracle Fusion General Ledger"
			else :
				product=""
			
			image_url=data[r][8]+".png"	
			print product
			
			
			try:
				elemuser=driver.find_element_by_id("sso_username")
				elemuser.send_keys(l_username)
				elempswd=driver.find_element_by_id("ssopassword")
				elempswd.send_keys(l_password)
				elemsubmit=driver.find_element_by_class_name("submit_btn")
				elemsubmit.click()
				element = WebDriverWait(driver, 100).until(
					EC.presence_of_element_located((By.ID,'pt1:panelCollection2:dialog3::ok'))
				)
			finally:
				driver.implicitly_wait(10)
				el = driver.find_element_by_id('pt1:panelCollection2:soc7::content')
				#print "EL:" + str(el.find_element_by_css_selector("*"))
				for option in el.find_elements_by_tag_name('option'):
					#print "Before if : Mathced " + option.text
					if option.text == data[r][0]:
						#print "After if : Matched " + option.text
						option.click()
						break
				
				#driver.implicitly_wait(20)
				
				prod1 = WebDriverWait(driver, 100).until(
					EC.presence_of_element_located((By.ID,'pt1:panelCollection2:soc6::content')))
				prod = driver.find_element_by_id('pt1:panelCollection2:soc6::content')	
				for option1 in prod.find_elements_by_tag_name('option'):
					print option1.text
					if option1.text == product:
						option1.click()
						break
				
				print product + "\t" + data[r][8] + "\n"
				driver.implicitly_wait(10)
				elembug=driver.find_element_by_id('pt1:panelCollection2:it15::content')
				elembug.send_keys(data[r][8]) ## from excel table
				
				
				          
				
				driver.implicitly_wait(10)
				driver.save_screenshot(image_url)	
				
				
				
				
				driver.implicitly_wait(5)
				elemsubmit=driver.find_element_by_id('pt1:panelCollection2:dialog3::ok')
				elemsubmit.click()
				driver.implicitly_wait(5)
			
			driver.close()
			
		r+=1
		
bug_associate_loop(1)	
