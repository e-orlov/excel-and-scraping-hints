Option Explicit

Sub Summe_markierter_Zellen()
Range("A1") = Application.WorksheetFunction.Sum(Selection)
End Sub