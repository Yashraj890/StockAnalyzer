import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook

class ExcelHelper:
    
    excelFilePath = ''
    sheetName = ''

    def __init__(self, filePath, sheetToRead):
        self.excelFilePath = filePath
        self.sheetName = sheetToRead

    def LoadExcelWorkBook(self):
        return load_workbook(self.excelFilePath)

    def GetExcelData(self):
        workbook = self.LoadExcelWorkBook()
        worksheet = workbook[self.sheetName]
        return worksheet.values
    
    def GetExcelDataColumns(self):
        data = self.GetExcelData()
        return next(data)[0:]

    def GetExcelDataFrame(self):
        data = self.GetExcelData()
        columns = self.GetExcelDataColumns()
        return pd.DataFrame(data, columns = columns)

