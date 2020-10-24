import glob
import os
import pandas as pd

InputFolder = 'IPsheet\\'
FileList = glob.glob(r'C:\Users\mhkas\PycharmProjects\make_pattern\IPsheet\*.xls')
print(FileList)
rb = pd.DataFrame(index=[])
for i in FileList:
    name = os.path.splitext(os.path.basename(i))[0]
    print(name)
    rb_tmp = pd.read_excel(InputFolder + name+'.xlsx', header=None , index_col=None)

    rb= pd.concat([rb,rb_tmp],axis=1)


    print(rb)

rb.to_excel(InputFolder + name+'.xls', header=None , index=None)
