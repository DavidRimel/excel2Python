import sys
import math
import numpy as np
from xlrd import open_workbook

global DEBUG
DEBUG = False

global CELL_EMPTY
CELL_EMPTY = u''

class excelSheet():

    def __init__(self, path2ExcelDoc, excelSheetNumber = 0):

        self.sheetNumb = excelSheetNumber
        self.workbook  = open_workbook(path2ExcelDoc)
        self.sheet     = self.workbook.sheets()[excelSheetNumber]
        self.cell      = self.sheet.cell
        self.nrows     = self.sheet.nrows
        self.ncols     = self.sheet.ncols


    def invalidExcelIndex(self, excelIndex):
         print ""
         print "================================"
         print ""
         print "Input = ", excelIndex
         print "This is not a valid excelIndex"
         print ""
         print "================================"
         print ""
         sys.exit(0)

    
    def splitExcelIndex( self, excelIndex ):

        col = excelIndex.translate(None,'0123456789')
        # checks that there is a letter and a number
        if ( col == excelIndex ) or ( col == '' ):
            self.invalidExcelIndex(excelIndex)

        row = int(excelIndex.translate(None, col))
        # checks that row is greater than or equal to one
        if ( row < 1 ):
            self.invalidExcelIndex(excelIndex)

        # checks for correct order "A1" valid and "1A" not valid
        if (( ord(list(excelIndex)[0]) - 64 ) < 1 ):
            self.invalidExcelIndex(excelIndex)

        return ( row, col )

    

    def excel2RowColIndex( self, excelIndex ):
        
        rowCol = self.splitExcelIndex( excelIndex )
        row = rowCol[0]
        col = rowCol[1]

        colList = []
        for char in list(col):
            colList.append(ord(char) - 64)

        mul = 1
        res = 0
        for pos in list(reversed(colList)):
            res += pos*mul
            mul *=26
    
        col = res
        return ( row-1, col-1  )



    def getCellValueByRowCol( self, row, col ):

        if ( row >= self.nrows ) or ( col >= self.ncols ):
            value = CELL_EMPTY
        else:    
            value = self.cell(row,col).value
        return value

    
    def getCellValueByExcelIndex( self, excelIndex ):

        rowCol = self.excel2RowColIndex( excelIndex )
        row = rowCol[0]
        col = rowCol[1]

        if ( row >= self.nrows ) or ( col >= self.ncols ):
            value = CELL_EMPTY
        else:    
            value = self.cell(row,col).value
        return value


    def getDataColFromExcel( self, excelIndex, debug = DEBUG):

        rowCol = self.excel2RowColIndex( excelIndex )
        col = rowCol[1]
        row = rowCol[0]
        rtnList = []
        value = self.getCellValueByRowCol( row, col )
        while (value != CELL_EMPTY):
            if debug:
                print "( row = ", row, ", col = ", col," ) = ", value
            rtnList.append(value)
            row = row + 1
            value = self.getCellValueByRowCol( row, col )

        #checks if rtnList is empty if so 
        # set it to None......
        # idk logically it makes more sense
        # that if you cant find the data 
        # the thing returned should be None.
        if not rtnList:
            rtnList = None

                  
        return np.array(rtnList)
       


