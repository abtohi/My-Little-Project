import pandas as pd
from openpyxl.workbook import Workbook
import os
import numpy as np

class Yuzu:
    def __init__(self,par_file,sheet):
        param         = pd.read_excel(par_file,sheet_name=sheet)
        param.columns = list(param.iloc[2,])
        act_type      = param.iloc[0,0]
        self.period   = param.iloc[0,1]
        param         = param[3:].dropna(axis=1, how='all')
        
        param.reset_index(inplace = True, drop = True)
        
        self.param      = param
        self.out_folder = 'Output'
        
        check_dir = os.path.basename(os.path.normpath(os.getcwd()))
        if(check_dir == 'Files'):
            self.file_list = os.listdir()
        else:
            os.chdir('Files')
            self.file_list = os.listdir()

        os.chdir('..')
        isExist = os.path.exists(self.out_folder)
        if isExist == False:
            os.mkdir(self.out_folder)
            
        
        if act_type == 'AVA SCORE':
            self.ava_score()
        elif act_type == 'AVA DISTRIBUSI':
            self.ava_distribusi()
        else:
            print('else')
    
    def ava_score(self):
        
        for file in self.file_list:
            wb = Workbook()
            re = file.replace('.xlsx','')
            wb.save(self.out_folder+'/'+re+' - AVA SCORE.xlsx')
            writer = pd.ExcelWriter(self.out_folder+'/'+re+' - AVA SCORE.xlsx', engine='xlsxwriter')

            for i, col in self.param.iterrows():
                #Data Preprocessing
                df         = pd.read_excel('Files/'+re+'.xlsx',sheet_name=self.param.loc[i,'sheets'], skiprows=3)[3:]
                df.reset_index(inplace = True, drop = True)

                df         = df[df['Unnamed: 0'].str.len() > 0]
                df         = df.dropna(axis=1, how='all')
                df.columns = df.columns.str.replace('\n','').str.replace('INNER',' INNER')
                len_rsa    = len(df.iloc[0,4:])
                n_rsa      = list(np.arange(4,4+len_rsa))
                rsa_ava    = df.iloc[:,[0] + n_rsa]

                rsa_ava    = rsa_ava.melt(id_vars='Unnamed: 0',var_name='kota',value_name='total')
                rsa_ava    = rsa_ava.rename(columns={'Unnamed: 0': 'provider'})
                ref        = pd.read_excel('Template/AVA SCORE.xlsx')
                ref        = ref.columns.str.strip()
                header     = pd.DataFrame(columns=ref)

                header.drop(
                    [
                        'Column 0',
                        'dashboard',
                        'Periode', 
                        'ROPERATOR AVA', 
                        'ROPERATOR new',
                        'AVA',
                        'AVA Score',
                        'AVA Score new',
                        'RSA AVA',
                        'Region AVA',
                        'Area AVA'
                    ], axis=1, inplace=True)

                header['ROPERATOR AVA'] = rsa_ava['provider']
                header['ROPERATOR new'] = rsa_ava['provider']
                header['AVA Score']     = rsa_ava['total']
                header['AVA Score new'] = rsa_ava['total']
                header['dashboard']     = 'AVA'
                header['Column 0']      = np.arange(1,len(rsa_ava)+1)
                header['Periode']       = self.period
                header['RSA AVA']       = rsa_ava['kota']
                header['Region AVA']    = df.columns[2]
                header['Area AVA']      = df.columns[3]

                if self.param.loc[i,'sheets']:
                    header['AVA'] = self.param.loc[i,'ava_type']

                header = header[list(ref)]          
                header.to_excel(writer, sheet_name=self.param.loc[i,'sheets'], index=False)

            writer.save()
            
    def ava_distribusi(self):
        for file in self.file_list:
            wb = Workbook()
            re = file.replace('.xlsx','')
            wb.save(self.out_folder+'/'+re+' - AVA DISTRIBUSI.xlsx')
            writer = pd.ExcelWriter(self.out_folder+'/'+re+' - AVA DISTRIBUSI.xlsx', engine='xlsxwriter')

            for i, col in self.param.iterrows():
                #Data Preprocessing
                df         = pd.read_excel('Files/'+re+'.xlsx',sheet_name=self.param.loc[i,'sheets'], skiprows=3)[3:]
                df.reset_index(inplace = True, drop = True)

                df         = df[df['Unnamed: 0'].str.len() > 0]
                df         = df.dropna(axis=1, how='all')
                df.columns = df.columns.str.replace('\n','').str.replace('INNER',' INNER')
                len_rsa    = len(df.iloc[0,4:])
                n_rsa      = list(np.arange(4,4+len_rsa))
                rsa_ava    = df.iloc[:,[0] + n_rsa]

                rsa_ava    = rsa_ava.melt(id_vars='Unnamed: 0',var_name='kota',value_name='total')
                rsa_ava    = rsa_ava.rename(columns={'Unnamed: 0': 'provider'})
                ref        = pd.read_excel('Template/AVA DISTRIBUSI.xlsx')
                ref        = ref.columns.str.strip()
                header     = pd.DataFrame(columns=ref)

                header.drop(
                    [
                        'Column 0',
                        'dashboard',
                        'Periode', 
                        'ROPERATOR AVA', 
                        'AVA',
                        'AVA Score',
                        'RSA AVA',
                        'Region AVA',
                        'Area AVA',
                        'Kategori'
                    ], axis=1, inplace=True)

                header['ROPERATOR AVA'] = rsa_ava['provider']
                header['AVA Score']     = rsa_ava['total']
                header['dashboard']     = 'AVA Distribusi'
                header['Column 0']      = np.arange(1,len(rsa_ava)+1)
                header['Periode']       = self.period
                header['RSA AVA']       = rsa_ava['kota']
                header['Region AVA']    = df.columns[2]
                header['Area AVA']      = df.columns[3]

                if self.param.loc[i,'sheets']:
                    header['AVA']      = self.param.loc[i,'ava_type']
                    header['Kategori'] = self.param.loc[i,'category']

                header = header[list(ref)]          
                header.to_excel(writer, sheet_name=self.param.loc[i,'sheets'], index=False)

            writer.save()

run_this = Yuzu('parameters.xlsx','PARAM')
