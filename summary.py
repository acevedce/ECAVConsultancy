import pandas as pd
import sys, os
import argparse

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
        print total
        for fac in self.cfile['Nro Factura'].to_string(index=False).split('\n'):
            amount = self.cfile.loc[self.cfile['Nro Factura'] == fac]['Monto'].to_string(index=False).split('\n')[-1]
            fsplit = fac.split('-')
            if len(fsplit)==3:
                self.cfactu.append({str(fsplit[2]):float(amount)})
        print self.cfactu
        for fac in self.bfile['Comentario'].to_string(index=False).split('\n'):
            self.bfactu.append(fac)
    
    def searchFactu(self):
        for f in self.bfactu:
            for fc in self.cfactu:
                if fc in f:
                    print self.bfile.loc[self.bfile['Comentario'] == f]['Monto']
                


if __name__ == '__main__':
    ll=[]

    opts = parse_args()
    pEx = parseExcel()
    pEx.loadExcel(opts.basefile, opts.crossfile)
    pEx.loadNroFactu()
    pEx.searchFactu()
    