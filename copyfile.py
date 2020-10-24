import pandas as pd
import os
import shutil
import openpyxl
InputFile = pd.read_excel('Inputlist.xlsx', header=0, index_col=0)
for i in InputFile.index:
    shutil.copy('format.xls', 'IPsheet\\format.xls')
    OutputFilePath = 'IPsheet\\'+'INITGT(ob7,ptw,' + str(InputFile.loc[i, '距離'])\
                     + ','+ str(InputFile.loc[i,'横位置']) + ').xls'
    os.rename('IPsheet\\format.xls', OutputFilePath)

    #print(i,InputFile.loc[i,'距離'])
    FileName_tmp = str(i)+'.xls'