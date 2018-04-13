#coding=utf-8
import xlrd
import xlsxwriter
import os
import threading
import re
import xlwt

class Bom(object):
    def __init__(self,model):
        self.i=0
        self.model=model
        self.th=[]
        self.rootdir=u'J:\\PIE Process Manual\\新工序文件(CMP)\\工序手冊\\{}\\'.format(model)
        self.re_=re.compile('[A-Z]{3}-[0-9]{3}-[0-9]{3}',re.S)
        self.workbook=xlwt.Workbook(encoding='gb18030')
        self.worksheet=self.workbook.add_sheet('sheet')
        self.get_file(self.rootdir)

    #返回excel对象
    def get_file(self,rootdir):
        files=os.listdir(rootdir)
        for fil in files:
            if os.path.splitext(fil)[1]!='.pdf':
                self.th.append(fil)
        return self.th

    #返回excel文件名和sheet表对象
    def openxls(self,fil):
        ar=[]
        xls=xlrd.open_workbook(self.rootdir+fil)
        xls_sheet_names=xls.sheet_names()
        for sheets in xls_sheet_names:
            if not 'ECN' in sheets:
                ar.append(xls.sheet_by_name(sheets))
        self.sheet(fil,ar)

    #单元格匹配
    def sheet(self,xls_book,xls_sheets):

        dic=set()
        for sht in xls_sheets:
            if 'ECN' not in sht.name:
                rows=sht.nrows
                for i in range(0,rows):
                    dic.add(sht.cell(i,9).value)
        self.parser(xls_book,dic)

    def parser(self,xls_book,dic):
        ar=[]
        for d in dic:
            try:
                key=re.findall(self.re_,d)
                if len(key)>0:
                    ar+=key
            except:
                pass

        #a=ar.sort()
        st='\n'.join(ar)
        self.wr(xls_book,st)

    def wr(self,xls_book,st):
        self.worksheet.write(self.i,0,xls_book[:-4])
        self.worksheet.write(self.i,1,st)
        self.i+=1

if __name__=='__main__':
    enter=raw_input('Enter model:>>>\n')
    f=Bom(enter)
    t=f.th
    for x in t:
        f.openxls(x)
    f.workbook.save('{}.xls'.format(enter))
