Sub SaveRowsAsTXT()
Dim wb As Excel.Workbook, wbNew As Excel.Workbook
Dim wsSource As Excel.Worksheet, wsTemp As Excel.Worksheet
Dim r As Long, c As Long
Dim filePath As String
Dim fileName As String
Dim rowRange As Range
Dim cell As Range

filePath = "C:\tmp\kdh\"

For Each cell In Range("A2", Range("A821").End(xlUp))
    Set rowRange = Range(cell.Address, Range(cell.Address).End(xlToRight))

    Set wsSource = ThisWorkbook.Worksheets("Sheet1")

    Application.DisplayAlerts = False 'will overwrite existing files without asking

    r = 1
    Do Until Len(Trim(wsSource.Cells(r, 1).Value)) = 0
        ThisWorkbook.Worksheets.Add ThisWorkbook.Worksheets(1)
        Set wsTemp = ThisWorkbook.Worksheets(1)

        For c = 2 To 16
            wsTemp.Cells((c - 1) * 2 - 1, 1).Value = wsSource.Cells(r, c).Value
        Next c
        fileName = filePath & wsSource.Cells(r, 1).Value

        wsTemp.Move
        Set wbNew = ActiveWorkbook
        Set wsTemp = wbNew.Worksheets(1)

        wbNew.SaveAs fileName & ".txt", xlTextWindows 'save as .txt
        wbNew.Close
        ThisWorkbook.Activate
        r = r + 1
    Loop

    Application.DisplayAlerts = True

Next
End Sub
