# coding=utf-8
import pandas as pd
import sys, os
import argparse
import re
import xlrd
import xlsxwriter

reload(sys)
sys.setdefaultencoding('utf8')


class parseExcel(object):
    def parse_args(self):
            parser = argparse.ArgumentParser()
            parser.add_argument("-b", "--basefile",required=True,help="root path to excel base file")
            parser.add_argument("-c", "--crossfile",required=True,help="root path to excel cross file")
            opts = parser.parse_args()
            return opts

    def loadExcel(self, basefile, crossfile):
        self.bfile=pd.read_excel(basefile)
        self.cfile=pd.read_excel(crossfile)
        return {'basefile':self.bfile, 'crossfile':self.cfile}
    
    def loadNroFactu(self):
        self.bfactu=[]
        self.cfactu=[]
        total=self.cfile.pivot_table(index=['Nro Factura'])
        totalbase=self.bfile.pivot_table(index=['Nro. de Deposito'])

        for fac in total.index:
            amount = total[(total.index.get_level_values('Nro Factura') == fac)]['Monto'].to_string(index=False).split('\n')[-1]
            deposit = total[(total.index.get_level_values('Nro Factura') == fac)]['Deposito Nro'].to_string(index=False).split('\n')[-1]
            fsplit = fac.split('-')
            if len(fsplit)==3:
                self.cfactu.append({str(fsplit[2]):{'Amount':float(amount),'Deposit':float(deposit)}})
        for fac in self.bfile['Comentario'].to_string(index=False).split('\n'):
            for unit in re.findall(r"\b\d+\b",fac):
                self.bfactu.append({str(unit):fac})

    def searchFactu(self):
        for f in self.bfactu:
            found = False
            for fc in self.cfactu:
                if f.keys()[0] == fc.keys()[0]:
                    print "found! - {}".format(self.bfile.loc[self.bfile['Comentario'] == f[str(f.keys()[0])].encode('utf-8')]['Monto'])
                    found = True
            if not found:
                print '{} - not founded'.format(f)

    def open_file(self, path):
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
    
        list=[]
        for row_num in range(1, sheet.nrows):
            row_value = sheet.row_values(row_num)
            if row_value[9] != '' and row_value[0]!= '':
       #         print row_value
                list.append(row_value)
        return list
    
    def files_diff(self, list0, list1):
        results={}
        noresults=[]
        for ii in list1:
            flag=False
            for i in list0:
                if ii[1].split('-')[2] in i[-3]:
                    if "{}-{}".format(ii[1],i[-3]) not in results.keys():
                        i.insert(0,True)
                        i.insert(0,ii[-3])
                        results["{}-{}".format(ii[1],i[-3])]=i
                    elif results["{}-{}".format(ii[1],i[-3])][-1] != ii[-3]:
                        i.insert(0,True)
                        i.insert(0,ii[-3])
                        results["{}-{}".format(ii[1],i[-3])]=i
                    flag=True
            if flag==False:
                if "{}-{}".format(ii[1],i[-3]) not in results.keys():
                    results["{}-{}".format(ii[1],i[-3])]=["none",ii[-3]]
        for i in list0:
            if i[1] != True and i[2]!='':
                noresults.append(i)

        return [results, noresults]

if __name__ == '__main__':
    ll=[]
    workbook = xlsxwriter.Workbook('Expenses01.xlsx')
    verificados = workbook.add_worksheet('Verificados')
    NoReportados = workbook.add_worksheet('NoReportados')
    NoVerificados = workbook.add_worksheet('NoVerificados')

    pEx = parseExcel()
    opts = pEx.parse_args()

    #pEx.loadExcel(opts.basefile, opts.crossfile)
    #pEx.loadNroFactu()
    #pEx.searchFactu()
    list0=pEx.open_file(opts.basefile)
    list1=pEx.open_file(opts.crossfile)
    dictionary, nodict = pEx.files_diff(list0, list1)
    for i in dictionary.keys():
        if dictionary[i][0] != "none":
            print "facturas: {} >> verificado :{} - reportado: {} = diferencia: {}".format(i, int(dictionary[i][-1]), int(dictionary[i][0]), int(dictionary[i][-1]) - int(dictionary[i][0]))
        else:
            print "factura no reportada: {}  con monto: {}".format(i, int(dictionary[i][1]))
    print "---------------------"
    for i in nodict:
        print "No tienen anotaciones de verificacion: {} -  con monto: {}".format(i[-3], int(i[-1]))
    
    verificados.write('A1','Factura')
    verificados.write('B1','Verificado')
    verificados.write('C1','Reportado')
    verificados.write('D1','Diferencia')
    row = 1
    for i in dictionary.keys():
        if dictionary[i][0] != "none":
            verificados.write(row, 0, i)
            verificados.write(row, 1, int(dictionary[i][-1]))
            verificados.write(row, 2, int(dictionary[i][0]))
            verificados.write(row, 3, int(dictionary[i][-1]) - int(dictionary[i][0]))
            row += 1

    NoReportados.write('A1','Factura')
    NoReportados.write('B1','Monto')
    row = 1
    for i in dictionary.keys():
        if dictionary[i][0] != "none":
            pass
        else:
            NoReportados.write(row, 0, i)
            NoReportados.write(row, 1, int(dictionary[i][1]))
            row += 1

    NoVerificados.write('A1','Factura')
    NoVerificados.write('B1','Monto')
    row = 1
    for i in nodict:
        NoVerificados.write(row, 0, i[-3])
        NoVerificados.write(row, 1, int(i[-1]))
        row += 1

    workbook.close()
        
        
        