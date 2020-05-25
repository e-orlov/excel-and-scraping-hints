Sub ZeilenLoeschen()
ActiveSheet.Rows("1:5").Delete
End Sub

Public Sub spalten_bereinigen()
With Worksheets(1).Range("1:1")
ActiveSheet.Rows(1).Find(what:="Search Engines", Lookat:=xlWhole).EntireColumn.Delete
ActiveSheet.Rows(1).Find(what:="Gruppen", Lookat:=xlWhole).EntireColumn.Delete
ActiveSheet.Rows(1).Find(what:="Diff", Lookat:=xlWhole).EntireColumn.Delete
ActiveSheet.Rows(1).Find(what:="Date", Lookat:=xlWhole).EntireColumn.Delete
End With

    'ActiveSheet.Columns().Delete
End Sub



Private Sub CommandButton1_Click()

Application.DisplayAlerts = False
' Namen der Tabellen anpassen
On Error Resume Next
ActiveWorkbook.Worksheets("All schedules").Delete
' ...
On Error GoTo 0
Application.DisplayAlerts = True

End Sub


Option Explicit

Sub ZeileLoeschen1()
   Dim LastRow As Long, LastCol As Integer
   Dim lRow As Long, lCol As Integer
   Dim i As Long
   
   'Maximale Werte feststellen
   With ActiveSheet
      lCol = .UsedRange.Columns.Count
      lRow = .UsedRange.Rows.Count
      For i = 1 To lCol
         LastRow = Application.WorksheetFunction.Max(.Cells(Rows.Count, i).End(xlUp).Row, LastRow)
      Next i
      For i = 1 To lRow
         LastCol = Application.WorksheetFunction.Max(.Cells(i, Columns.Count).End(xlToLeft).Column, LastCol)
      Next i
   End With
   
   'letzte Zeile mit Daten komplett löschen
   Rows(LastRow).EntireRow.Delete
End Sub

Option Explicit

Sub ZeileLoeschen2()
   Dim LastRow As Long, LastCol As Integer
   Dim lRow As Long, lCol As Integer
   Dim i As Long
   
   'Maximale Werte feststellen
   With ActiveSheet
      lCol = .UsedRange.Columns.Count
      lRow = .UsedRange.Rows.Count
      For i = 1 To lCol
         LastRow = Application.WorksheetFunction.Max(.Cells(Rows.Count, i).End(xlUp).Row, LastRow)
      Next i
      For i = 1 To lRow
         LastCol = Application.WorksheetFunction.Max(.Cells(i, Columns.Count).End(xlToLeft).Column, LastCol)
      Next i
   End With
   
   'letzte Zeile mit Daten komplett löschen
   Rows(LastRow).EntireRow.Delete
End Sub

Option Explicit

Sub ZeileLoeschen3()
   Dim LastRow As Long, LastCol As Integer
   Dim lRow As Long, lCol As Integer
   Dim i As Long
   
   'Maximale Werte feststellen
   With ActiveSheet
      lCol = .UsedRange.Columns.Count
      lRow = .UsedRange.Rows.Count
      For i = 1 To lCol
         LastRow = Application.WorksheetFunction.Max(.Cells(Rows.Count, i).End(xlUp).Row, LastRow)
      Next i
      For i = 1 To lRow
         LastCol = Application.WorksheetFunction.Max(.Cells(i, Columns.Count).End(xlToLeft).Column, LastCol)
      Next i
   End With
   
   'letzte Zeile mit Daten komplett löschen
   Rows(LastRow).EntireRow.Delete
End Sub

*********************************************

Public Sub DateiWaehlen()
    
    Dim vntFile As Variant
    Dim pfad As String
    Dim objWorkbook As Workbook
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        .Title = "!!! DATEI WÄHLEN !!!"
        .AllowMultiSelect = True
        .InitialFileName = pfad
        
        With .Filters
            If .Count > 0 Then Call .Clear
            .Add "Excel 2010", "*.xlsx"
            .Add "Excel 2003", "*.xls"
            .Add "Alle", "*.*"
        End With
        
        If .Show = -1 Then
            For Each vntFile In .SelectedItems
                Set objWorkbook = Workbooks.Open(vntFile)
                
ActiveSheet.Rows("1:5").Delete 'Erste 5 Zeilen löschen

With Worksheets(1).Range("1:1") 'Bereich auswählen
ActiveSheet.Rows(1).Find(what:="Search Engines", Lookat:=xlWhole).EntireColumn.Delete 'Kolumne "Search Engines" löschen
ActiveSheet.Rows(1).Find(what:="Diff", Lookat:=xlWhole).EntireColumn.Delete 'Kolumne "Diff" löschen
ActiveSheet.Rows(1).Find(what:="Date", Lookat:=xlWhole).EntireColumn.Delete 'Kolumne "Date" löschen
End With 'Unnötige Kolumnen gelöscht

Application.DisplayAlerts = False 'Löschung des unnötigen Tabs
On Error Resume Next
ActiveWorkbook.Worksheets("All schedules").Delete
On Error GoTo 0
Application.DisplayAlerts = True ' Tab gelöscht

' letzte Zeile löschen
   Dim LastRow1 As Long, LastCol1 As Integer
   Dim lRow1 As Long, lCol1 As Integer
   Dim i As Long
   
   With ActiveSheet
      lCol1 = .UsedRange.Columns.Count
      lRow1 = .UsedRange.Rows.Count
      For i = 1 To lCol1
         LastRow1 = Application.WorksheetFunction.Max(.Cells(Rows.Count, i).End(xlUp).Row, LastRow1)
      Next i
      For i = 1 To lRow1
         LastCol1 = Application.WorksheetFunction.Max(.Cells(i, Columns.Count).End(xlToLeft).Column, LastCol1)
      Next i
   End With
   
   Rows(LastRow1).EntireRow.Delete
' letzte Zeile gelöscht

' letzte Zeile löschen
   Dim LastRow2 As Long, LastCol2 As Integer
   Dim lRow2 As Long, lCol2 As Integer
   Dim a As Long
   
   With ActiveSheet
      lCol2 = .UsedRange.Columns.Count
      lRow2 = .UsedRange.Rows.Count
      For a = 1 To lCol2
         LastRow2 = Application.WorksheetFunction.Max(.Cells(Rows.Count, a).End(xlUp).Row, LastRow2)
      Next a
      For a = 1 To lRow2
         LastCol2 = Application.WorksheetFunction.Max(.Cells(a, Columns.Count).End(xlToLeft).Column, LastCol2)
      Next a
   End With
   
   Rows(LastRow2).EntireRow.Delete
' letzte Zeile gelöscht

' letzte Zeile löschen
   Dim LastRow3 As Long, LastCol3 As Integer
   Dim lRow3 As Long, lCol3 As Integer
   Dim o As Long
   
   With ActiveSheet
      lCol3 = .UsedRange.Columns.Count
      lRow3 = .UsedRange.Rows.Count
      For o = 1 To lCol3
         LastRow3 = Application.WorksheetFunction.Max(.Cells(Rows.Count, o).End(xlUp).Row, LastRow3)
      Next o
      For o = 1 To lRow3
         LastCol3 = Application.WorksheetFunction.Max(.Cells(o, Columns.Count).End(xlToLeft).Column, LastCol3)
      Next o
   End With
   
   Rows(LastRow3).EntireRow.Delete
' letzte Zeile gelöscht

                Call objWorkbook.Save
                Set objWorkbook = Nothing
            Next
        Else
            
            MsgBox "!!! KEINE DATEI AUSGEWÄHLT !!!", vbExclamation, "!!! WARNUNG !!!"
            
        End If
    End With
    
End Sub

