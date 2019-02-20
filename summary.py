# coding=utf-8
import pandas as pd
import sys, os
import argparse
import re

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

if __name__ == '__main__':
    ll=[]

    opts = parse_args()
    pEx = parseExcel()
    pEx.loadExcel(opts.basefile, opts.crossfile)
    pEx.loadNroFactu()
    pEx.searchFactu()
    