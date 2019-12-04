import xlrd
import pygeoip
import time
import xlsxwriter
#Insert Your File Path Here
loc = 'E:/PycharmProjects/Clicker/Clicker/upload_Download_2019120104021245/Report.xlsx'
workbook2 = xlrd.open_workbook(loc)
sheet = workbook2.sheet_by_index(0)
row = 0
column = 0
totalRow = 0
temp = []
sheet.cell_value(0, 0)
#Give The Column Number In Which Your Ip Exist
for i in range(sheet.nrows):
    print(sheet.cell_value(i, 1))
    temp.append(sheet.cell_value(i, 1))
    totalRow += 1
#Enter Your Address For New Sheet
workbook2 = xlsxwriter.Workbook('E:/PycharmProjects/Clicker/Clicker/upload_Download_2019120104021224/REPORTResult.xlsx')
worksheet = workbook2.add_worksheet()
#Choose From Where You Want To Enter Data In The Sheet
row = 0
column = 0
for i in range(totalRow):
    print(i)
    print(temp[i])
    tempstr = temp[i]
    try:
        gip = pygeoip.GeoIP('GeoLiteCity.dat')
        res = gip.record_by_addr(tempstr)
        time.sleep(0.1) #Sleep Required
        tempcity = res["city"] #Getting City Out Of It
        print(tempcity)
        worksheet.write(row, column, tempcity) #Enter The Data Into The New Sheet
        worksheet.write(row, 1, temp[i]) #Enter The Data Into The New Sheet
    except:
        pass

    row += 1#Increment Rows
workbook2.close()#Save The Sheet