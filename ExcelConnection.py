import openpyxl as xlr
from Singleton import Singleton

class ExcelConnection(metaclass = Singleton):

    def connectToExcel(self):
        return xlr.Workbook()
