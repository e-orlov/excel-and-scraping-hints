Option Explicit

Sub x()
   Dim i As Long, j As Integer
   
   j = Cells.SpecialCells(xlCellTypeLastCell).Column
   For i = Cells.SpecialCells(xlCellTypeLastCell).Row To 1 Step -1
      If Application.CountIf(Cells(i, 1).Resize(, j), "*Suchwort*") < 1 Then Rows(i).Delete
   Next
End Sub 