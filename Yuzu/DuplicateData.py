import os
import pandas as pd
import datetime
from openpyxl.workbook import Workbook

class DuplicateYuzu:
    def __init__(self,trans_sheet, trans_id, gu_file,gu_sheet,gu_id,nu_dup):
        self.trans_sheet = trans_sheet
        self.trans_id    = trans_id
        self.gu_file     = gu_file
        self.gu_sheet    = gu_sheet
        self.gu_id       = gu_id
        self.nu_dup      = nu_dup
        self.out_folder  = 'Output'
        
        check_dir = os.path.basename(os.path.normpath(os.getcwd()))
        if(check_dir == 'Files to Duplicate'):
            self.file_list = os.listdir()
        else:
            os.chdir('Files to Duplicate')
            self.file_list = os.listdir()

        os.chdir('..')
        isExist = os.path.exists(self.out_folder)
        if isExist == False:
            os.mkdir(self.out_folder)
            
    def duplicate(self):
        for file in self.file_list:
            wb = Workbook()
            re = file.replace('.xlsx','')
            wb.save(self.out_folder+'/'+re+' - Result.xlsx')
            writer = pd.ExcelWriter(self.out_folder+'/'+re+' - Result.xlsx', engine='xlsxwriter')

            print("Start :",'{:%Y-%b-%d %H:%M:%S}'.format(datetime.datetime.now()))

            transaksi = pd.read_excel('Files to Duplicate/'+re+'.xlsx', sheet_name=self.trans_sheet)
            jml = pd.read_excel(self.gu_file,sheet_name=self.gu_sheet, skiprows = range(0, 3))

            jml = jml[jml[self.nu_dup].notnull()]
            jml = jml.round({self.nu_dup:0})
            jml[self.nu_dup] = jml[self.nu_dup].astype(int)

            print("Read Data Finished :",'{:%Y-%b-%d %H:%M:%S}'.format(datetime.datetime.now()))

            data = pd.DataFrame([], columns = []) 
            for i in transaksi.index:
                n = False
                for j in jml.index:
                    if transaksi[self.trans_id][i] == jml[self.gu_id][j]:
                        data = data.append(pd.concat([transaksi.iloc[[i],]]*(jml[self.nu_dup][j])), ignore_index=True)
                        n = True
                if n == False:
                    data = data.append(transaksi.iloc[[i],], ignore_index=True)

            print("Looping Finished :",'{:%Y-%b-%d %H:%M:%S}'.format(datetime.datetime.now()))

            data.to_excel(writer,sheet_name='Output',index=False)
        writer.save()

        print("Finish",'{:%Y-%b-%d %H:%M:%S}'.format(datetime.datetime.now()))
        
execute = DuplicateYuzu('SPLIT DATA','QSbjNum_Rekrut','Yellow Yuzu April 2022 - Data Duplikasi to DP.xlsx','Sheet1','RESPID RECRUIT','DUPLICATE GUIDANCE')
execute.duplicate()
