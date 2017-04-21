import excel2Python

excelPath = 'exampleSheet.xlsx'

workbook = excel2Python.excelSheet( excelPath )

Data = workbook.getDataColFromExcel( 'A1' )

print Data

