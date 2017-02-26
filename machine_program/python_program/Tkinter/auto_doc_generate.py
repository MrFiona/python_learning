#!/usr/bin/env python 

"""
-----------------------------------------------------------------------------------
2015.07.16   Grey
***********************************************************************************

Use the win32com to process word document and generate a set of ENW documents.
"""
import win32com,os,re,time,datetime,string
from win32com.client import Dispatch,constants

def enw_documents_generate(document_name):
	enw_document_type                   = raw_input("Please input the type for ENW Doc (E/e for electronic and others for paper):").upper()
	pwd                                 = os.getcwd()
	document_list                       = []
	hex_list                            = []
	pack_flag_for_replace               = 'pack_name'
	word_app                            = win32com.client.Dispatch('Word.Application')
	word_app.Visible                    = 0
	word_app.DisplayAlerts              = 0
	document_name                       = pwd + '\\' + document_name
	base_time                           = time.strftime('%Y_%m_%d',time.localtime(time.time()))
	date_today                          = datetime.datetime(string.atoi(base_time[0:4]),string.atoi(base_time[5:7]),string.atoi(base_time[8:10]))
	date_three_days_later               = date_today + datetime.timedelta(days = 85)
	date_today_str                      = re.sub(r' 00.00.00','',str(date_today))
	date_three_days_later_str           = re.sub(r' 00.00.00','',str(date_three_days_later))
	date_today_str                      = re.sub(r'(\d\d\d\d)-(\d\d)-(\d\d)',r'\2/\3/\1',date_today_str)
	date_three_days_later_str           = re.sub(r'(\d\d\d\d)-(\d\d)-(\d\d)',r'\2/\3/\1',date_three_days_later_str)
	#document                            = word_app.Documents.Open(FileName = document_name)
	#word_app.Selection.Find.ClearFormatting()  
	#word_app.Selection.Find.Replacement.ClearFormatting()
	for root,dirs,files in os.walk('.'):
		hex_list               = [f for f in os.listdir('.') if f.endswith('.rar')]
	for hex_file in hex_list:
		pack_name = re.sub('.rar','',hex_file)
		document                            = word_app.Documents.Open(FileName = document_name)
		word_app.Selection.Find.ClearFormatting()  
		word_app.Selection.Find.Replacement.ClearFormatting()
		word_app.Selection.Find.Execute('pack_name', False, False, False, False, False, True, 1, True, pack_name, 2)
		word_app.Selection.Find.Execute('today', False, False, False, False, False, True, 1, True, date_today_str, 2)
		if enw_document_type == 'E':
			word_app.Selection.Find.Execute('Three_days_later', False, False, False, False, False, True, 1, True, '', 2)
		else :
			word_app.Selection.Find.Execute('Three_days_later', False, False, False, False, False, True, 1, True, date_three_days_later_str, 2)
		document.SaveAs(pwd + '\\' + 'CTENW_' + pack_name +'.docx')
	#document.Close()
	word_app.Documents.Close() 
	word_app.Quit() 

enw_documents_generate('ENW_temp.docx')
print "All the ENW documents has been generated!"
