 For Each ws In ActiveWorkbook.Worksheets  
   > chandoo.org/forum/threads/vba-loop-to-activate-sheets-in-order
last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
   > Stack Overflow 
open_price = Cells(2, column + 2).Value
   > Stack Overflow 
 If Cells(i + 1, column).Value <> Cells(i, column).Value Then
 ticker = Cells(i, column).Value
 Cells(row, column + 8).Value = ticker    
    > Study Group
Cells(row, column + 10).NumberFormat = "0.00%"
    > wallstreetmojo.com/vba-format-number/
volume = volume + Cells(i, column + 6).Value
    > Xpert Learning Assistant
 If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row)) Then
    > Study Group  
