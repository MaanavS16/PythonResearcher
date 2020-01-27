from google import google
import time
import xlrd
import xlwt

book = xlwt.Workbook()
sheet = book.add_sheet('Terms')
workbook = xlrd.open_workbook('C:/Users/MaanavSingh/dev/PCInput/SearchTerms.xlsx')
worksheet = workbook.sheet_by_name('Terms')
#print(worksheet.cell(0,0).value)
arrTerms = []
#Read and store terms in final revision xlsx
for i in range(worksheet.nrows):
    time.sleep(.05)
    if worksheet.cell(i, 0) == xlrd.empty_cell.value:
        #print("break")
        break
    else:
        term = worksheet.cell(i, 0).value
        arrTerms.append(term)
        sheet.write(i,0,term)
        #print(worksheet.cell(i, 0).value)
#print(arrTerms)

for i in range(len(arrTerms)):
    counter = 0
    query = arrTerms[i]
    search_results = google.search(query, 1)
    for result in search_results:
        counter += 1
        if counter == 1:
            sheet.write(i,2,result.description)
            print(arrTerms[i],result.description)
'''
for i in range(len(arrTerms)):
    query = arrTerms[i] + " AP Human Geography"
    for j in search(query,tld="com", lang='en', num=10, stop=1, pause=1.5):
        sheet.write(i,2,j)
        print(j)
'''
book.save('C:/Users/AnilSingh/dev/PCInput/SearchTerms1.xls')
print('saved')
