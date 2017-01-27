from google import search
import xlrd, xlwt
import time

book_read = xlrd.open_workbook("queries.xls")         ## give your excel sheet name(this program and your sheet should be 
sheet_read = book_read.sheet_by_index(0)             ## at same location or you will have to change location constantly)
index = sheet_read.nrows
l=[0 for i in range(index)]
search_data = [[sheet_read.cell_value(r,0).encode('ascii','replace')] for r in range(index)]  ## replacing utf encoded words with ascii filler

book_write = xlwt.Workbook()
sheet_write = book_write.add_sheet("Sheet", cell_overwrite_ok=True)

ctr = -1
for i in range(index/40+1):
    for j in range(40):
        ctr += 1
        if ctr < index:
            for url in search(search_data[ctr][0], tld='com', lang='en', stop=1):
                print url
                l[ctr]=str(url)
                sheet_write.write(ctr, 0, l[ctr])
                break
        else:
            break
    book_write.save('link.xls')
    if ctr < index:
        time.sleep(30*60)        
        

## Due to google service policy more than 45 (or so) automated queries it will
## Temporarily block your IP for another automted query
##Thus give about a 30 minutes time delay after every 40 queries
