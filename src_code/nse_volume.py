# import time
from nsetools import Nse
import easygui as gui
import xlrd
import xlsxwriter

# get the input file with nse code
def nse_data():
	gui.msgbox(msg="Select input xl file with NSE codes\nRefer 'stock_watchlist.xlsx file'", title="Input file", ok_button='Browse')
	ip_file = gui.fileopenbox(msg=None, title=None, default='*', filetypes=None, multiple=False)

	# ip_file = "D:\python\workspace\stock_volume\stock_watchlist.xlsx"

	nse = Nse()
		
	#op result file creation
	workbook = xlsxwriter.Workbook("NSE_trading_data_output.xlsx")
	op_sheet = workbook.add_worksheet("NSE_data")

		 # Add a bold format to use to highlight cells.
	bold = workbook.add_format({'bold': 1})
	date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

	op_sheet.write(0,0,"NSE code",bold)
	op_sheet.write(0,1,"Date",bold)
	op_sheet.write(0,2,"Delivery QTY",bold)
	op_sheet.write(0,3,"Volume",bold)
	op_sheet.write(0,4,"Delivery percentage",bold)


		

	# print(ip_file)
	wb = xlrd.open_workbook(ip_file) 
	sheet = wb.sheet_by_index(0) 
	  
	# For row 0 and column 0 
	no_of_stocks = sheet.nrows
	# print(no_of_stocks)
	# print(sheet.cell_value(0, 0)) 

	for row_no in range(1,no_of_stocks):
		nse_code = str((sheet.cell_value(row_no, 0)))
		print(nse_code)
		stock = nse.get_quote(nse_code)
		
		deliv_percnt = 	stock['deliveryToTradedQuantity']
		deliv = 	stock['deliveryQuantity']
		volume = 	stock['quantityTraded']
		date = 	stock['secDate']
		
		#----  write data into a xl file ----
		
		op_sheet.write(row_no,0,nse_code)	
		op_sheet.write(row_no,1,date)	
		op_sheet.write(row_no,2,deliv)	
		op_sheet.write(row_no,3,volume)	
		op_sheet.write(row_no,4,deliv_percnt)	
		
		
			
		
	print("DONE!!")	
	workbook.close()
	gui.msgbox(msg="SUCCESS!!! \n Check Output file 'NSE_trading_data_output.xlsx' in application directory", title="DONE", ok_button='FINISH')	



	
#----------------------------------main----------------------------------------

if __name__=="__main__"	:
	print("main starts\n")
	try:
		nse_data()
	except:
		gui.exceptionbox(msg="ERROR!!!!\nCheck if you have kept NSE_trading_data_output file open", title="Error")
	
	