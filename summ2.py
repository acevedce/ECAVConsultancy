import xlrd

def open_file(path):
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    list=[]
    for row_num in range(1, sheet.nrows):
        row_value = sheet.row_values(row_num)
        if row_value[9] != '' and row_value[0]!= '':
   #         print row_value
            list.append(row_value)
    return list

if __name__ == "__main__":
    path = "Planilla excel H.N S.A.-/FEBRERO 2016.-/ALCOSUR_022016.xlsx"
    list0=open_file(path)
    path = "HN SA 2016 ( DEPOSITOS )/DEPOSITOS ENERO 2016.-/ALCOSUR.xlsx"
    list1=open_file(path)
    
    for ii in list1:
        flag=False
        for i in list0:
            if ii[1].split('-')[2] in i[-3]:
                print "{} - found with key {}".format(i, ii[1])
                flag=True
        if flag==False:
            print "NONE RESULT FOR: {}".format(ii[1])