import pandas as pd
import os
import shutil
import openpyxl

#InputFile = pd.read_csv('Inputlist.csv', header=0, index_col=0, encoding='SHIFT-JIS')
InputFile = pd.read_excel('Inputlist.xlsx', header=0, index_col=0)

#print(InputFile.index)

#with pd.ExcelWriter('') as writer:

"""
for i in InputFile.index:
    shutil.copy('format.xls', 'IPsheet\\format.xls')
    OutputFilePath = 'IPsheet\\'+'INITGT(ob7,ptw,' + str(InputFile.loc[i, '距離'])\
                     + ','+ str(InputFile.loc[i,'横位置']) + ').xls'
    os.rename('IPsheet\\format.xls', OutputFilePath)

    #print(i,InputFile.loc[i,'距離'])
    FileName_tmp = str(i)+'.xls'

"""
for i in InputFile.index:
    OutputFilePath = 'IPsheet\\'+'INITGT(ob7,ptw,' + str(InputFile.loc[i, '距離'])\
                     + ','+ str(InputFile.loc[i,'横位置']) + ').xlsx'
    book = openpyxl.load_workbook(OutputFilePath)
    sheet = book['Sheet1']
    sheet['C8']=InputFile.loc[i, '距離']
    book.save(OutputFilePath)


FormatFile = pd.read_excel('format.xls')
print(FormatFile)



