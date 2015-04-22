#coding=utf-8
import xlrd, xlwt

TXT_FILE_PATH = '/Users/yangshaoli/work/zicaihui/usage_statistics/clicknum.txt'
EXCEL_FILE_PATH = '/Users/yangshaoli/work/zicaihui/usage_statistics/clicknum.xls'

if __name__ == "__main__":
    f = open(TXT_FILE_PATH,'r')
    result = list()

    for line in f.readlines():
        result.append(line)

    wb = xlwt.Workbook(encoding="utf-8")
    table = wb.add_sheet('sheet1')
    row = 0
    style0 = xlwt.XFStyle()
	
    for line in result:
	    items = line.split(' ')

	    table.write(row, 0, items[0], style0)
	    print items[0]
	    
	    if len(items) == 3:
	        table.write(row, 1, items[1], style0)
	        print items[1]
	        
	        ctype = 2
	        table.write(row, 2, int(items[2][:-1]))
	        print items[2][:-1]
	    row += 1
	
    wb.save(EXCEL_FILE_PATH)