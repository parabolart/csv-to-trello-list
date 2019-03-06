import xlwings as xw
import time, datetime, csv

sheet = xw.books.active.sheets.active
values = {}
values[0] = {
    'status': '',
    'name': '',
    'leadsource': '',
    'bid': '',
    'contract': '',
    'deposit': '',
    'ptd': '',
    'co': '',
    'balance': '',
    'criticaldate': '',
    'client': '',
    'contact': '',
    'contactnumber': '',
    'drawings': ''
}
i = 0

for row in range(13, 180):
    #print(row)
    for col in range(8,9):
        #print(col)
        if sheet.range((row,col)).value == "Production":
            i += 1
            values[i] = values[0].copy()
            values[i]['status'] = sheet.range((row,col)).value
            values[i]['name'] = sheet.range((row,col-1)).value
            if sheet.range((row,col+1)).value != None:
                values[i]['leadsource'] = sheet.range((row,col+1)).value
            if sheet.range((row,col+2)).value != None:
                values[i]['bid'] = sheet.range((row,col+2)).value
            if sheet.range((row,col+4)).value != None:
                values[i]['contract'] = sheet.range((row,col+4)).value
            if sheet.range((row,col+5)).value != None:
                values[i]['deposit'] = sheet.range((row,col+5)).value
            if sheet.range((row,col+6)).value != None:
                values[i]['ptd'] = sheet.range((row,col+6)).value
            if sheet.range((row,col+7)).value != None:
                values[i]['co'] = sheet.range((row,col+7)).value
            if sheet.range((row,col+8)).value != None:
                values[i]['balance'] = sheet.range((row,col+8)).value
            if sheet.range((row,col+10)).value != None:
                values[i]['criticaldate'] = sheet.range((row,col+10)).value.strftime("%B") #date() works
            if sheet.range((row,col+11)).value != None:
                values[i]['client'] = sheet.range((row,col+11)).value
            if sheet.range((row,col+12)).value != None:
                values[i]['contact'] = sheet.range((row,col+12)).value
            if sheet.range((row,col+13)).value != None:
                values[i]['contactnumber'] = sheet.range((row,col+13)).value
            if sheet.range((row,col+16)).value != None:
                values[i]['drawings'] = sheet.range((row,col+16)).value
            with open('190305_production.csv', mode='w') as csvfile:
                csvfile_writer = csv.writer(csvfile,delimiter='|', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                csvfile_writer.writerow([values[i]['name'],
                    values[i]['leadsource'],
                    values[i]['client'],
                    values[i]['contact'],
                    str(values[i]['bid']),
                    str(values[i]['ptd']),
                    str(values[i]['balance'])])
            '''
            print(values[i]['name']
                + '\n   lead: ' + values[i]['leadsource']
                + '\n   bid: ' + values[i]['bid']
                + '\n   contract: ' + str(values[i]['contract'])
                + '\n   deposit: ' + str(values[i]['deposit'])
                + '\n   ptd: ' + str(values[i]['ptd'])
                + '\n   co: ' + str(values[i]['co'])
                + '\n   balance: ' + str(values[i]['balance'])
                + '\n   critical date: ' + str(values[i]['criticaldate'])
                + '\n   client: ' + str(values[i]['client'])
                + '\n   contact: ' + str(values[i]['contact'])
                + '\n   #: ' + str(values[i]['contactnumber'])
                + '\n   drawings: ' + values[i]['drawings'])'''
            #print("The Row is: "+str(row)+" and the column is "+str(col))

'''
x = sheet['A3'].value

while True:
    y = sheet['A3'].value
    if x != y:
        x = y
        print (x)
    time.sleep(0.5)
'''
csvfile.close()
