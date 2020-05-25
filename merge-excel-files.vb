Sub Combine()
    Dim J As Integer
    Dim s As Worksheet

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    For Each s In ActiveWorkbook.Sheets
        If s.Name <> "Combined" Then
            Application.GoTo Sheets(s.Name).[a1]
            Selection.CurrentRegion.Select
            ' Don't copy the headings
            Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
            Selection.Copy Destination:=Sheets("Combined"). _
              Cells(Rows.Count, 1).End(xlUp)(2)
        End If
    Next
End Sub