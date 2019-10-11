from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from win32com.client.gencache import EnsureDispatch as Dispatch
from paramiko_expect import SSHClientInteraction
from datetime import date, timedelta, datetime
from pdfminer.converter import TextConverter
from win32com.client import constants
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams
import win32com.client as win32
from cStringIO import StringIO
from progress.bar import Bar
from shutil import copyfile
import win32com.client
from datetime import *
from ftplib import *
import time as tym
import subprocess
import pyautogui
import pdfminer
import paramiko
import getpass
import socket
import sys
import os
import re

#from datetime import timedelta
#from datetime import date
#import datetime

#to-do: account for 7 digit member numbers in case when CHECK.RELN doesn't find a loan for the 6 digit member number

#REGARDING DATES: THIS WILL PULL LOANS FROM ONBASE THAT WERE FROM LAST WEEK (no matter what day this is being run) AS DEFINED IN THE FUNCTION check_the_date_of_report AND IT WILL PULL LOANS FROM EMAILS DEFINED BY THE GLOBAL VARIABLES start & end WHICH MAY BE DIFFERENT DATES. 

#A NOTE ABOUT EMAILS: THE FUNCTION printSubjectAndCreationTime IS SUPPOSED TO PULL THE NUMBER OF EMAILS THAT IS PASSED INTO THIS FUNCTION. HOWEVER DUE TO AN APPARENT BUG?? IN THE win32com.client MODULE, ONLY ABOUT 700-800 EMAILS ARE PULLED FROM THE HELPDESK FOLDER, NO MATTER WHAT NUMBER IS PASSED INTO THE FUNCTION. BUT THAT SHOULD BE ENOUGH FOR OUR PURPOSES.

#RUN AFTER NATHAN IMPORTS PREVIOUS DAY'S/WEEK'S FILES INTO ONBASE IN THE MORNING ~7:30AM

#check for proper monitor resolution

if socket.gethostname() == 'MB-TASK-SYS':
	program_paths_file = raw_input("Enter file with program paths, e.g. C:\Users\sys-task\Desktop\Python Programs\program_paths.py : ") or "C:\Users\sys-task\Desktop\Python Programs\program_paths.py"
else:
	program_paths_file = raw_input("Enter file with program paths, e.g. H:\PersonalSave\Desktop\scripts\python\program_paths.py : ") or "H:\PersonalSave\Desktop\scripts\python\program_paths.py"

try:
	execfile(program_paths_file)
except:
	pass

if pyautogui.size() != (1920, 1080):
	exit()

write_to_file = 1
	
#workbook = raw_input("Enter location of spreadsheet file containing passwords, e.g. H:\\PersonalSave\\Desktop\\scripts\\python\\readPasswordProtectedExcel\\logins.xlsx: ") or 'H:\\PersonalSave\\Desktop\\scripts\\python\\readPasswordProtectedExcel\\logins.xlsx'
password = getpass.getpass("Enter master password: ")
xlApp = win32com.client.Dispatch("Excel.Application")
xlwb = xlApp.Workbooks.Open(passwordsFile, True, False, None, password)
xlws = xlwb.Sheets(1) # counts from 1, not from 0
server 				= str(xlws.Cells(2,1))
server_user 		= str(xlws.Cells(3,1))
server_pass 		= str(xlws.Cells(3,2))
server_user2 		= str(xlws.Cells(5,1))
server_pass2 		= str(xlws.Cells(5,2))
ultrafis_pass 		= str(xlws.Cells(6,1))[:-2]
mortgage_user 		= str(xlws.Cells(1,1))
mortgage_pass 		= str(xlws.Cells(1,2))
#tcl_pass 			= str(xlws.Cells(7,1))[:-2]
xlApp.Quit()

tcl_pass = getpass.getpass("Enter TCL password: ")
print server, server_user, server_pass, server_user2, server_pass2, ultrafis_pass, tcl_pass, mortgage_user, mortgage_pass

def progbar(message, seconds):
	bar = Bar('%40s' % message, max=seconds)
	for i in range(seconds):
		tym.sleep(1)
		bar.next()
	bar.finish()	

try:
	os.remove('output.txt')
	os.remove('checkrelnLOANS.txt')
except:
	pass

	
#start = date.today() - timedelta(1)
#end = date.today() + timedelta(1)

#DEFAULT TIME RANGE
start = date.today() - timedelta(8)
end = date.today() + timedelta(1)


#now = datetime.now()
#beginning_of_last_week = now-timedelta(days=8) #final version will be 8
#end_of_last_week = now-timedelta(days=0) #final version will be 0

list_of_five_digit_numbers_in_email = []
list_of_six_digit_numbers_in_email = []
list_of_seven_digit_numbers_in_email = []
list_of_eight_digit_numbers_in_email = []
list_of_nine_digit_numbers_in_email = []
list_of_member_numbers = []


root = "\\\\onbase\\data$\\REPORTS\\"
tempdir = "X:\\IT\\Private\\temp\\Andrey\\FICS Loan Coupons\\"

#function to convert a pdf file into text, or extract text out of a pdf file, however you want to phrase it
def pdf_to_text(pdf_file_full_path):
    # PDFMiner boilerplate
    rsrcmgr = PDFResourceManager()
    sio = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, sio, codec=codec, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Extract text
    fp = open(pdf_file_full_path, 'rb')
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
    fp.close()

    # Get text from StringIO
    text = sio.getvalue()

    # Cleanup
    device.close()
    sio.close()

    return text

def check_if_new_loan_summary_report(textfile):
	with open(textfile) as f:
		for line_num, line_text in enumerate(f):
			if "New Loans Entered Summary" in line_text:
				return True, line_num + 2 # The date of the pdf is stored 2 lines after "New Loans Entered Summary"
				f.close()
	f.close()
	return False, 0

def check_the_date_of_report(textfile, num):
	with open(textfile) as f:
		for line_num, line_text in enumerate(f):
			if line_num == num:
				for check_every_day in range(2+datetime.today().weekday(),8+datetime.today().weekday()): #datetime.today().weekday() is in case this is run on a day that is not Monday
					if ' ' + (datetime.today() - timedelta(days=check_every_day)).isoformat()[8:10] + ',' in line_text:
						print "GOGOGOGOGOGOGOGOGGO"
						f.close()
						return True
				f.close()
				os.remove(textfile)
				return False

def get_loan_numbers(textfile):
	with open(textfile) as f:
		for line_text in f:
			if re.search("(\d{7})\D",line_text):
				if re.search("(\d{8})\D",line_text):
					if re.search("(\d{9})\D",line_text):
						if re.search("(\d{10})\D",line_text):
							print line_text
							loans.append(line_text)
							continue
						print line_text
						loans.append(line_text)
						continue
					print line_text
					loans.append(line_text)
					continue
				print line_text
				loans.append(line_text)
		f.close()
		return

i = 0
loans = []
for root, dirs, files in os.walk(root):
	for fname in files:
		i = i + 1
		if fname.endswith('.pdf'):
			fullpath = os.path.join(root, fname)
			file_stat = os.stat(fullpath)
			file_modification_time = datetime.fromtimestamp(file_stat.st_mtime)
			if file_modification_time.date() > start and file_modification_time.date() < end and int(os.path.getsize(fullpath)) < 50000:
				print('%s modified %s'%(fullpath, file_modification_time))
				converted_text_file = os.path.join(tempdir, fname[:-3] + 'txt')
				open_converted_text_file = open(converted_text_file, 'wb')
				try:
					open_converted_text_file.write(pdf_to_text(fullpath))
				except:
					print "LMAO IT APPEARS PDFMINER FAILED CONVERTING THIS PDF FILE TO TEXT. IT IS PROBABLY NOT A LOAN COUPON FILE"
				open_converted_text_file.close()
				boo, num = check_if_new_loan_summary_report(converted_text_file)
				if not boo:
					print "OK OK OK"
					os.remove(converted_text_file)
				if boo:
					if check_the_date_of_report(converted_text_file, num):
						get_loan_numbers(converted_text_file)

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
class Oli():
	def __init__(self, outlook_object):
		self._obj = outlook_object

	def items(self):
		array_size = self._obj.Count
		for item_index in xrange(1,array_size):
			yield (item_index, self._obj[item_index])						

if write_to_file == 1:
	old_stdout = sys.stdout
	output = open("output.txt", "w")
	sys.stdout = output			

def printSubjectAndCreationTime(subfolder, num):
	list_of_six_digit_numbers_in_email = []
	list_of_seven_digit_numbers_in_email = []
	list_of_eight_digit_numbers_in_email = []
	list_of_nine_digit_numbers_in_email = []
	member_is_in_subject = 0
	nine_in_subject_flag = 0
	eight_in_subject_flag = 0
	seven_in_subject_flag = 0
	six_in_subject_flag = 0
	five_in_subject_flag = 0
	for inx, subsubfolder in Oli(folder.Folders).items():
		if (subsubfolder.Name == "Inbox"):
			print folder.Name
			print subsubfolder.Name
			folder_of_interest = subsubfolder.Items
			message = folder_of_interest.GetLast()
			message_body = message.Body
			message_subject = message.Subject
			ct = message.CreationTime
			tym.sleep(1)
			print start
			print end
			print date(int("20" + str(ct)[6:8]),int(str(ct)[0:2]), int(str(ct)[3:5]))
			print ct
			#if (datetime.date(int("20" + str(ct)[6:8]),int(str(ct)[0:2]), int(str(ct)[3:5])) > start) and (datetime.date(int("20" + str(ct)[6:8]),int(str(ct)[0:2]), int(str(ct)[3:5])) < end):
			line = str(message_subject)[0:60], str(ct)[0:8]
				
			message_subject = str(line)

			#if ("oupon" in message_subject) or ("oupon" in message_body) and ("Hold Digest" not in str(line)) and (re.search("(\d{6})\D",message_subject) or (re.search("(\d{6})\D",message_body))):
			#if (("oupon" in message_subject) or ("oupon" in message_body) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{6})\D",message_subject) or re.search("(\d{6})\D",message_body)):
			#if (("oupon" in message_subject) or ("oupon" in message_body) or ("ayment" in message_subject) or ("ayment" in message_body) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{6})\D",message_subject) or re.search("(\d{6})\D",message_body)):
			if (("oupon" in message_subject) or ("oupon" in message_body) or (("ayment" in message_subject) or ("ayment" in message_body) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject))) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject) or ("oan" in message_body) or ("oan" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{5})\D",message_subject) or re.search("(\d{5})\D",message_body)):
			#if (("Book" in message_subject) or ("book" in message_subject)) and (re.search("(\d{5})\D",message_subject) or re.search("(\d{5})\D",message_body)):
				if ("Hold Digest" not in str(line)) and ("SWBC FTP Server Notification" not in str(line)) and ("Completed WO" not in str(line)):
					print '%80s' % str(line)
					member_number_raw = re.compile("(\d{5})\D")
					print member_number_raw.findall(message_subject)
					print member_number_raw.findall(message_body)
						
				if ("Service Request" in message_subject):
					for k in range(0, len(member_number_raw.findall(message_body))):
						#list_of_member_numbers.append(member_number_raw.findall(message_body)[k])
						print member_number_raw.findall(message_body)[k]
					
						
			for i in range (0, num):
				message = folder_of_interest.GetPrevious()
				message_body = message.Body
				message_subject = message.Subject
				ct = message.CreationTime
				member_is_in_subject = 0
				eight_in_subject_flag = 0
				seven_in_subject_flag = 0

				if (date(int("20" + str(ct)[6:8]),int(str(ct)[0:2]), int(str(ct)[3:5])) > start) and (date(int("20" + str(ct)[6:8]),int(str(ct)[0:2]), int(str(ct)[3:5])) < end):
					try:
						line = str(message_subject)[0:60], str(ct)[0:8]
					except:
						#print "exception" + message_subject
						continue
					
					message_subject = str(line)
			
					if ("Hold Digest" in str(line)) or ("SWBC FTP Server Notification" in str(line)) or ("Completed WO" in str(line)):
						continue
					#if (("oupon" in message_subject) or ("oupon" in message_body) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{6})\D",message_subject) or re.search("(\d{6})\D",message_body)):
					#if (("oupon" in message_subject) or ("oupon" in message_body) or ("ayment" in message_subject) or ("ayment" in message_body) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{6})\D",message_subject) or re.search("(\d{6})\D",message_body)):
					#if (("oupon" in message_subject) or ("oupon" in message_body) or (("ayment" in message_subject) or ("ayment" in message_body) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject))) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{6})\D",message_subject) or re.search("(\d{6})\D",message_body)):
					if (("oupon" in message_subject) or ("oupon" in message_body) or ((("ayment" in message_body) or ("ayment" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject))) or (((("RELN" in message_body) or ("RELN" in message_subject)) or ("ortgage" in message_body) or ("ortgage" in message_subject) or ("oan" in message_body) or ("oan" in message_subject)) and (("book" in message_body) or ("Book" in message_body) or ("book" in message_subject) or ("Book" in message_subject)))) and (re.search("(\d{5})\D",message_subject) or re.search("(\d{5})\D",message_body)):
						print '%80s' % str(line)
						
						print date(int("20" + str(ct)[6:8]),int(str(ct)[0:2]), int(str(ct)[3:5]))
						print ct
						
						member_number_raw = re.compile("(\d{5})\D")
						if ("Service Request" in message_subject):
							for k in range(0, len(member_number_raw.findall(message_body))):
								#list_of_member_numbers.append(member_number_raw.findall(message_body)[k])
								print member_number_raw.findall(message_body)[k]
						
						member_number_raw9 = re.compile("(\d{9})\D")
						if member_number_raw9.findall(message_subject):
							list_of_nine_digit_numbers_in_email.append(member_number_raw9.findall(message_subject))
							member_is_in_subject = 1
							nine_in_subject_flag = 1
						if member_number_raw9.findall(message_body):
							list_of_nine_digit_numbers_in_email.append(member_number_raw9.findall(message_body))
						
						member_number_raw8 = re.compile("(\d{8})\D")
						if member_number_raw8.findall(message_subject):
							list_of_eight_digit_numbers_in_email.append(member_number_raw8.findall(message_subject))
							member_is_in_subject = 1
							eight_in_subject_flag = 1
						if member_number_raw8.findall(message_body):
							list_of_eight_digit_numbers_in_email.append(member_number_raw8.findall(message_body))
						
						member_number_raw7 = re.compile("(\d{7})\D")
						if member_number_raw7.findall(message_subject):
							list_of_seven_digit_numbers_in_email.append(member_number_raw7.findall(message_subject))
							member_is_in_subject = 1
							seven_in_subject_flag = 0
						if member_number_raw7.findall(message_body):
							list_of_seven_digit_numbers_in_email.append(member_number_raw7.findall(message_body))
										
						member_number_raw6 = re.compile("(\d{6})\D")
						if member_number_raw6.findall(message_subject):
							list_of_six_digit_numbers_in_email.append(member_number_raw6.findall(message_subject))
							member_is_in_subject = 1
							six_in_subject_flag = 1
						if member_number_raw6.findall(message_body):
							list_of_six_digit_numbers_in_email.append(member_number_raw6.findall(message_body))

						member_number_raw5 = re.compile("(\d{5})\D")
						if member_number_raw5.findall(message_subject):
							list_of_five_digit_numbers_in_email.append(member_number_raw5.findall(message_subject))
							member_is_in_subject = 1
						if member_number_raw5.findall(message_body):
							list_of_five_digit_numbers_in_email.append(member_number_raw5.findall(message_body))	
						
						if member_is_in_subject == 1:
						
							if (list_of_nine_digit_numbers_in_email) and (nine_in_subject_flag == 1):
							
								print "member_is_in_subject = 1. list_of_nine_digit_numbers_in_email[0]. nine_in_subject:"
							
								print (str(list_of_nine_digit_numbers_in_email[0]))[2:8]
								#print (str(list_of_nine_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_nine_digit_numbers_in_email[0]))[2:8])
						
							elif (list_of_eight_digit_numbers_in_email) and (eight_in_subject_flag == 1):
							
								print "member_is_in_subject = 1. list_of_eight_digit_numbers_in_email[0]. eight_in_subject:"
							
								print (str(list_of_eight_digit_numbers_in_email[0]))[2:8]
								#print (str(list_of_eight_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_eight_digit_numbers_in_email[0]))[2:8])
							elif list_of_seven_digit_numbers_in_email and (seven_in_subject_flag == 1):
							
								print "member_is_in_subject = 1. list_of_seven_digit_numbers_in_email[0]. seven_in_subject:"
								
								print (str(list_of_seven_digit_numbers_in_email[0]))[2:8]
								list_of_member_numbers.append((str(list_of_seven_digit_numbers_in_email[0]))[2:8])
							elif list_of_six_digit_numbers_in_email and (six_in_subject_flag == 1):
							
								print "member_is_in_subject = 1. list_of_six_digit_numbers_in_email[0]. six_in_subject:"
								
								print (str(list_of_six_digit_numbers_in_email[0]))[2:8]
								list_of_member_numbers.append((str(list_of_six_digit_numbers_in_email[0]))[2:8])
							else:
								#if list_of_five_digit_numbers_in_email:
								
								print "member_is_in_subject = 1. list_of_five_digit_numbers_in_email[0]:"
								
								print (str(list_of_five_digit_numbers_in_email[0]))[2:7]
								list_of_member_numbers.append((str(list_of_five_digit_numbers_in_email[0]))[2:7])
						
						else:
						
							if list_of_nine_digit_numbers_in_email:
								
								print "member_is_in_subject = 0. list_of_nine_digit_numbers_in_email[0]:"
								
								print (str(list_of_nine_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_nine_digit_numbers_in_email[0]))[3:9])
						
							elif list_of_eight_digit_numbers_in_email:
								
								print "member_is_in_subject = 0. list_of_eight_digit_numbers_in_email[0]:"
								
								print (str(list_of_eight_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_eight_digit_numbers_in_email[0]))[3:9])
							elif list_of_seven_digit_numbers_in_email:
							
								print "member_is_in_subject = 0. list_of_seven_digit_numbers_in_email[0]:"
							
								print (str(list_of_seven_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_seven_digit_numbers_in_email[0]))[3:9])
							elif list_of_six_digit_numbers_in_email:
							
								print "member_is_in_subject = 0. list_of_six_digit_numbers_in_email[0]:"
							
								print (str(list_of_six_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_six_digit_numbers_in_email[0]))[3:9])
							else:
							
								print "member_is_in_subject = 0. list_of_five_digit_numbers_in_email[0]:"
							
								print (str(list_of_five_digit_numbers_in_email[0]))[3:8]
								list_of_member_numbers.append((str(list_of_five_digit_numbers_in_email[0]))[3:8])
						
						
						if list_of_nine_digit_numbers_in_email:
							print "LINE 232"
							print list_of_nine_digit_numbers_in_email[0]
							if member_is_in_subject == 0:
								print (str(list_of_nine_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_nine_digit_numbers_in_email[0]))[3:9])
							else:
								print (str(list_of_nine_digit_numbers_in_email[0]))[2:8]
								list_of_member_numbers.append((str(list_of_nine_digit_numbers_in_email[0]))[2:8])
						
						if list_of_eight_digit_numbers_in_email:
							print "LINE 242"
							print list_of_eight_digit_numbers_in_email[0]
							if member_is_in_subject == 0:
								print (str(list_of_eight_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_eight_digit_numbers_in_email[0]))[3:9])
							else:
								print (str(list_of_eight_digit_numbers_in_email[0]))[2:8]
								list_of_member_numbers.append((str(list_of_eight_digit_numbers_in_email[0]))[2:8])
						
						if list_of_seven_digit_numbers_in_email:
							print "LINE 252"
							print list_of_seven_digit_numbers_in_email[0]
							if member_is_in_subject == 0:
								print (str(list_of_seven_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_seven_digit_numbers_in_email[0]))[3:9])
							else:
								print (str(list_of_seven_digit_numbers_in_email[0]))[2:8]
								list_of_member_numbers.append((str(list_of_seven_digit_numbers_in_email[0]))[2:8])		
								
						if list_of_six_digit_numbers_in_email:
							print "LINE 262"
							print list_of_six_digit_numbers_in_email[0]
							if member_is_in_subject == 0:
								print (str(list_of_six_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_six_digit_numbers_in_email[0]))[3:9])
							else:
								print (str(list_of_six_digit_numbers_in_email[0]))[2:8]
								list_of_member_numbers.append((str(list_of_six_digit_numbers_in_email[0]))[2:8])
								
						if list_of_five_digit_numbers_in_email:
							print "LINE 262"
							print list_of_five_digit_numbers_in_email[0]
							if member_is_in_subject == 0:
								print (str(list_of_five_digit_numbers_in_email[0]))[3:9]
								list_of_member_numbers.append((str(list_of_five_digit_numbers_in_email[0]))[3:9])
							else:
								print (str(list_of_five_digit_numbers_in_email[0]))[2:7]
								list_of_member_numbers.append((str(list_of_five_digit_numbers_in_email[0]))[2:7])
						
				else:
					continue
				
				list_of_five_digit_numbers_in_email = []
				list_of_six_digit_numbers_in_email = []
				list_of_seven_digit_numbers_in_email = []
				list_of_eight_digit_numbers_in_email = []
				list_of_nine_digit_numbers_in_email = []

for inx, folder in Oli(mapi.Folders).items():
	if (folder.Name == "HelpDesk"):
		printSubjectAndCreationTime(folder, 1000)

if write_to_file == 1:
	sys.stdout = old_stdout
	output.close()
	ftpsession = FTP('vcu', server_user, server_pass)
	todir='/data/AMFCU/STAGING/RELN/'
	ftpsession.cwd(todir)

	fromdir='H:\\PersonalSave\\Desktop\\scripts\\python\\FICSOrderingRELNCoupons\\'
	from_filename = os.path.join(fromdir, 'output.txt')

	file_object = open(from_filename, 'rb')
	ftpsession.storbinary('STOR '+ 'output.txt', file_object)
	file_object.close()

	ftpsession.quit()
	
prompt = "$"
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(server, username=server_user, password=server_pass)
interact = SSHClientInteraction(ssh, timeout=1200, newline ="\r", display=True, encoding='utf-8')
interact.send('/data/AMFCU/STAGING/RELN/masterScript.sh %s %s %s %s %s' % (server, server_user2, server_pass2, ultrafis_pass, tcl_pass))

tym.sleep(60)

def checkSize():
	print "IN THE checkSize FUNCTION"
	stdin,stdout,stderr = ssh.exec_command("ls -ltr `find /data/AMFCU/STAGING/RELN/checkreln.txt` | awk '{print $5}'")
	checkrelnSize1 = str(stdout.readlines())
	tym.sleep(60)
	stdin,stdout,stderr = ssh.exec_command("ls -ltr `find /data/AMFCU/STAGING/RELN/checkreln.txt` | awk '{print $5}'")
	checkrelnSize2 = str(stdout.readlines())
	return checkrelnSize1, checkrelnSize2

for i in range(0,60):
	checkrelnSize1, checkrelnSize2 = checkSize()
	if checkrelnSize1 != checkrelnSize2:
		#time.sleep(60)
		progbar("still working...", 60)
	else:
		break

#interact.send('/data/AMFCU/STAGING/RELN/finalBattle.sh')
#ssh.exec_command('/data/AMFCU/STAGING/RELN/finalBattle.sh')

stdin,stdout,stderr = ssh.exec_command("ls -ltr `find /data/AMFCU/STAGING/RELN/checkrelnLOANS.txt` | awk '{print $6, $7, $8}'")
fileWithLoanNumbers1 = str(stdout.readlines())
print "/data/AMFCU/_HOLD_/STAGING/RELN/checkrelnLOANS.txt file creation date: " + fileWithLoanNumbers1

stdin,stdout,stderr = ssh.exec_command("ls -ltr `find /data/AMFCU/STAGING/RELN/checkrelnLOANS.txt` | awk '{print $5}'")
fileWithLoanNumbers2 = str(stdout.readlines())
print "/data/AMFCU/_HOLD_/STAGING/RELN/checkrelnLOANS.txt file size: \t\t" + fileWithLoanNumbers2

#interact.expect(prompt)
interact.send('exit')

fromdir='/data/AMFCU/STAGING/RELN/'
todir='H:\\PersonalSave\\Desktop\\scripts\\python\\FICSOrderingRELNCoupons\\'
ftp = FTP('vcu')
ftp.login(user=server_user, passwd=server_pass)
to_filename = os.path.join(todir, 'checkrelnLOANS.txt')
ftp.cwd(fromdir)
file_object = open(to_filename, 'wb')
ftp.retrbinary('RETR '+ 'checkrelnLOANS.txt', file_object.write)
file_object.close()
ftp.quit()

#for line in open('checkrelnLOANS.txt'):
#    loans.append(line.rstrip().split(','))
		
with open('checkrelnLOANS.txt') as f:
	for line_text in f:
		print line_text
		loans.append(line_text)
print loans								
		
#print finalized list of loans from Onbase and emails
for j in loans:
	print j[0:9]
#Open Mortgage Servicer
#pyautogui.click(20,1050)
#pyautogui.hotkey("ctrl","esc")
#tym.sleep(1)
#pyautogui.typewrite("mortgage")
#tym.sleep(1)
#pyautogui.press("enter")
#tym.sleep(5)
try:
	os.startfile(MS_location)
except:
	try:
		pyautogui.click(20,1050)
	except:
		exit()
	tym.sleep(1)
	try:
		pyautogui.typewrite("mortgage")
	except:
		exit()
	tym.sleep(1)
	try:
		pyautogui.press("enter")
	except:
		exit()

#tym.sleep(12)
tym.sleep(30)

try:
	pyautogui.typewrite(mortgage_user)
except:
	exit()
try:
	pyautogui.press('tab')
except:
	exit()
try:
	pyautogui.typewrite(mortgage_pass)
except:
	exit()
try:
	pyautogui.press('enter')
except:
	exit()
tym.sleep(5)

#collapse all folders
try:
	pyautogui.click(74, 193)
except:
	exit()
try:
	pyautogui.press('down')
except:
	exit()
try:
	pyautogui.press('up')
except:
	exit()
try:
	pyautogui.click(74, 193)
except:
	exit()
tym.sleep(1)

#click "Billing"
try:
	pyautogui.click(50, 240)
except:
	exit()
tym.sleep(2)

#click "Automated Coupons"
try:
	pyautogui.doubleClick(50, 280)
except:
	exit()
tym.sleep(2)

#Start New MSCOUPON.FIL File checkbox
try:
	pyautogui.click(614, 226)
except:
	exit()
tym.sleep(2)

#Enter ABA & Branch numbers
try:
	pyautogui.click(480, 276)
except:
	exit()
tym.sleep(2)
try:
	pyautogui.typewrite("314977133")
except:
	exit()
try:
	pyautogui.click(542, 297)
except:
	exit()
tym.sleep(2)
try:
	pyautogui.typewrite("1")
except:
	exit()

#Include Coupon Data checkbox
try:
	pyautogui.click(612, 482)
except:
	exit()
tym.sleep(2)

#Next Payment Due Date Plus 12
try:
	pyautogui.click(897, 352)
except:
	exit()
tym.sleep(2)
try:
	pyautogui.typewrite("12")
except:
	exit()

#click "Loan Select"
try:
	pyautogui.click(321, 537)
except:
	exit()
tym.sleep(2)

#input loan numbers
for j in loans:
	try:
		pyautogui.typewrite(j[0:9])
	except:
		exit()
	tym.sleep(2)
	try:
		pyautogui.press('enter')
	except:
		exit()
	tym.sleep(2)
	for letter in "dre":
		pyautogui.press('tab')
	try:
		pyautogui.press('enter')
	except:
		exit()
	tym.sleep(2)
	try:
		pyautogui.doubleClick(633,241)
	except:
		exit()
	tym.sleep(2)

pyautogui.alert(text='Check loan numbers before continuing.', title='Check loan numbers before continuing.', button='Check loan numbers before continuing.')	

pyautogui.hotkey('alt', 'e')
tym.sleep(5)
pyautogui.click(917, 979)
tym.sleep(20)
pyautogui.click(1154, 532)
tym.sleep(5)

os.rename ('\\\\ficsapp\\FICS\mscoupon.fil', 'X:\\IT\\Private\\Mailers & Coupons\\RE-LOAN-' + date.today().isoformat()[5:7] + date.today().isoformat()[8:10])

#subprocess.call([SZ_location, 'a', 'pv3l0c1ty', '-y', 'X:\\IT\\Private\\Mailers & Coupons\\RE-LOAN-' + date.today().isoformat()[5:7] + date.today().isoformat()[8:10] + '.zip'] + ['X:\\IT\\Private\\Mailers & Coupons\\RE-LOAN-' + date.today().isoformat()[5:7] + date.today().isoformat()[8:10]])

raw_input("Press Enter to continue...")