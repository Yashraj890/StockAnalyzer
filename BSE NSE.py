import pandas as pd
from ExcelHelper import ExcelHelper 

NSEFilePath = r'D:\\NSEData\\NSE.xlsx'
BSEFilePath = r'D:\\NSEData\\BSE.xlsx'

excelHelperNSE = ExcelHelper(NSEFilePath, 'NSE')
nsewb = excelHelperNSE.LoadExcelWorkBook()
nsedf = excelHelperNSE.GetExcelDataFrame()
nsecolumn = excelHelperNSE.GetExcelDataColumns()

excelHelperBSE = ExcelHelper(BSEFilePath, 'Equity')
bsedf = excelHelperBSE.GetExcelDataFrame()

nsenewws = nsewb.create_sheet('nsenew')
nsenewws.title = 'nsenew'
nsenewws.append(nsecolumn)
for secid in nsedf['Security Id'].unique():
    filterdf = bsedf.query('`Security Id` == @secid')
    if not filterdf.empty:
        finalfilterdf = nsedf.query('`Security Id` == @secid')
        nsenewws.append([finalfilterdf['ISIN No'].values[0],finalfilterdf['Security Id'].values[0],finalfilterdf['Security Name'].values[0],filterdf['Industry'].values[0], filterdf['Group'].values[0][0]])
nsewb.save(NSEFilePath)
