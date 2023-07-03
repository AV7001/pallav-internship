import os
import pandas as pd
import openpyxl
import tqdm
from art import *

class Excel_Adjuster:
    def single_file_conversion(self,path):
        print(f'Path Found: {path}')
        path=path.replace('\\','/')
        name=path.split('/')[-1]
        o_path=path.replace(name,'')
        try:
            data=pd.read_excel(path)
            op_data=pd.DataFrame(columns=data.columns)
            #display(op_data.columns)
            sites=sorted(list(set(data['SiteID'].to_list())))
            for indx,val in enumerate(sites):
                classed_data=data.loc[data['SiteID']==val]
                op_data=pd.concat([op_data,classed_data],ignore_index=True)
                s = pd.Series([None])
                op_data=pd.concat([op_data,s])

            op_data=op_data[['Site','SiteID','Name','OccuredOn','ClearedOn']]
            os.remove(os.path.join(o_path,name))
            print(f'{name} conversion Complete !')
            op_data.to_excel(os.path.join(o_path,name),index=False)
        except:
            print('Already Done on this File Cannot do it again')

    def multiple_file_conversion(self,folder_path):
        for file in tqdm.tqdm(os.listdir(folder_path)):
            if file.endswith('.xlsx'):
                self.single_file_conversion(os.path.join(folder_path,file))
        print('All file Converted !!')
        


    def __init__(self):
        while True:
            os.system('cls')
            main_logo=text2art('Excel Convertor')
            print(main_logo)
            try:
                choice=int(input('\nUser Input \n 1 for Single File Conversion \n 2 for multiple File Conversion \n 3 to quit \n User:> '))
            except:
                choice=1000

            if choice==1:
                single_path=input('Enter path of Single File: ')
                self.single_file_conversion(single_path)
                continue_flow=input('Press Any key to continue')
            elif choice==2:
                path_folder=input('Enter path of the folder having multiple files: ')
                self.multiple_file_conversion(path_folder)
                continue_flow=input('Press Any key to continue')
            elif choice==3:
                break
            else:
                print('Invalid Please Enter Again')


Excel_Adjuster()