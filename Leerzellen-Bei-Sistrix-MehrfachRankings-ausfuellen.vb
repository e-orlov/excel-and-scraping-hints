Sub ausfuellen()
With Intersect(Columns("A:A"), ActiveSheet.UsedRange)
.SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
.Value = .Value
End With
End Sub