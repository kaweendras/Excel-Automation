path="C:\Users\salit\Desktop\VBS\Test.xlsx"

set xlapp=CreateObject("Excel.Application")
xlapp.DisplayAlerts=False

set xlwb=xlapp.WorkBooks.open(path)

set sht1=xlwb.Sheets(1)

introws = sht1.UsedRange.Rows.Count
intcols = sht1.UsedRange.Columns.Count

MsgBox sht1.Cells(1,intcols)
MsgBox introws 
MsgBox intcols 

'close the file
