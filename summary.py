# coding=utf-8
import pandas as pd
import sys, os
import argparse
import re
import xlrd

reload(sys)
sys.setdefaultencoding('utf8')

def parse_args():
        parser = argparse.ArgumentParser()
        parser.add_argument("-b", "--basefile",required=True,help="root path to excel base file")
        parser.add_argument("-c", "--crossfile",required=True,help="root path to excel cross file")
        opts = parser.parse_args()
        return opts

class parseExcel(object):
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
        for ii in list1:
            flag=False
            for i in list0:
                if ii[1].split('-')[2] in i[-3]:
                    i.insert(0,str(ii[1]))
                    print "{} - found with key".format(i)
    #                list.append(i)
                    flag=True
            if flag==False:
                print "No results for factu num: {}".format(ii[1])

if __name__ == '__main__':
    ll=[]

    opts = parse_args()
    pEx = parseExcel()
    pEx.loadExcel(opts.basefile, opts.crossfile)
    pEx.loadNroFactu()
    pEx.searchFactu()
    list0=pEx.open_file(opts.basefile)
    list1=pEx.open_file(opts.crossfile)
    print pEx.files_diff(list0, list1)
        