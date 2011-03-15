import xlrd
## initial output file 
ru = open('/home/willy/data/ru','r+')
fr = open('/home/willy/data/fr','r+')
de = open('/home/willy/data/de','r+')
it = open('/home/willy/data/it','r+')
es = open('/home/willy/data/es','r+')
pt = open('/home/willy/data/pt','r+')
el = open('/home/willy/data/else','r+')

## write to specipic file 
def write2file(ftype,eng,local):
        text=eng+'^'+local
	if ftype==2:
		output=fr
	elif ftype==3:
		output=ru
	elif ftype==4:
		output=es
	elif ftype==5:
		output=de
	elif ftype==6:
		output=pt
	elif ftype==7:
		output=it
        else: 
		output=el 
	output.write(" ".join(text.split('\n'))+'\n')
## start 
book = xlrd.open_workbook('/home/willy/data/refineby.xls',encoding_override='utf-8')
#print book.nsheets
#print book.sheet_names()
sheetnum = book.nsheets
#for s in range(sheetnum)
for s in range(sheetnum):
    sh = book.sheet_by_index(s)
    ##column number row number
    rownum = sh.nrows
    print rownum
    colnum = sh.ncols
    #for i in range(1,rownum):
    for i in range(1,rownum):
        for j in range(1,colnum):
	    value=unicode(sh.cell_value(i,j)).strip()
	    if j==1:
	        eng = value.encode('utf-8')
	    else:
	        write2file(j,eng,value.encode('utf-8'))
ru.close()
fr.close()
de.close()
it.close()
es.close()
pt.close()
el.close()
