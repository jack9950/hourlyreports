import win32com.client as win32

def convert_file(file_name):

	#fname = 'C:\\Users\\jackson.ndiho.IQOR\\Documents\\script\\report.xls'
	fname = file_name
	excel = win32.gencache.EnsureDispatch('Excel.Application')
	wb = excel.Workbooks.Open(fname)

	wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
	wb.Close()                               #FileFormat = 56 is for .xls extension
	excel.Application.Quit()