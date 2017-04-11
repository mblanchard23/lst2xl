from openpyxl import Workbook
from openpyxl.cell import cell

def lst2xl(lst,xlws,start_cell='A1'):
	

	row_count = len(lst)
	
	if not row_count: # Do nothing if data set is empty
		return None


	col_count = len(lst[0])
	

	row_start = xlws[start_cell].row 
	col_start = cell.column_index_from_string(xlws[start_cell].column)


	coords = {'x':0,'y':0}
	
	for row in range(row_start,row_start+row_count):
		for col in range(col_start,col_start+col_count):
			try:
				xlws.cell(row=row,column=col,value=lst[coords['y']][coords['x']])
			except UnicodeDecodeError:
				pass
			

			coords['x'] += 1
		coords['y'] += 1
		coords['x'] = 0

	return xlws


def lst2wb(lst,fpath=None):
	if fpath:
		wb_name = fpath	
	else:
		wb_name = raw_input("Please enter a workbook name or fpath: ")
		if wb_name[-5:] != '.xlsx':
			wb_name += '.xlsx'
	wb = Workbook()
	ws = wb.active
	lst2xl(lst,ws)
	try:
		wb.save(wb_name)
		print "Successfully saved to %s" % (wb_name)
	except:
		print "Failed to save - check you have write access"
