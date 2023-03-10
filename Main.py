from pathlib import Path
import xlwings as xw
import time

import shutil,os


SOURCE_DIR ='Alltypes'
SOURCE_DIF='NotAplicable'
SOURCE_PROC='Processed'


#Reading Files all types
allFiles= os.listdir(SOURCE_DIR)
for diferente in allFiles:
    #if files are diferente than excel docs move
    if diferente.endswith('.xlsx')!=True:
        print(diferente)
        shutil.move(SOURCE_DIR+'/'+diferente,SOURCE_DIF)
    else:
        shutil.move(SOURCE_DIR+'/'+diferente,SOURCE_PROC)





#Getting Files with xlsx extensions
excel_files= list(Path(SOURCE_PROC).glob('*.xlsx'))
combined_wb= xw.Book()
tm= time.localtime()
tmstamp= time.strftime('%Y-%m-%d_%H%M',tm)

for excel_file in excel_files:
    wb=xw.Book(excel_file)
    for sheet in wb.sheets:
        sheet.api.Copy(After= combined_wb.sheets[0].api)

wb.close()

combined_wb.sheets[0].delete()
combined_wb.save(f'Master{tmstamp}.xlsx')

if len(combined_wb.app.books)==1:
    combined_wb.app.quit()
else: 
    combined_wb.close()



    
