from openpyxl import utils

def lst2xl(lst,xlws,start_cell):
	

	row_count = len(lst)
	
	if not row_count: # Do nothing if data set is empty
		return None


	col_count = len(lst[0])
	

	row_start = ws[start_cell].row 
	col_start = column_index_from_string(ws[start_cell].column)


	coords = {'x':0,'y':0}
	
	for row in range(row_start,row_start+row_count):
		for col in range(col_start,col_start+col_count):

			xlws.cell(row=row,column=col,value=lst[coords['y']][coords['x']])
			coords['x'] += 1
		coords['y'] += 1
		coords['x'] = 0

#
	wb.save(save_dir+ rev_filename)
	return rev_filename


