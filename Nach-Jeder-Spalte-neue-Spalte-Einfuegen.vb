Sub JedeZweiteSpalte()
    Dim i As Integer

    For i = ActiveSheet.UsedRange.Columns.Count To 2 Step -1
        Columns(i).Insert
    Next i
End Sub