Sub main()
    Dim isht As Long
    Dim allSht As Worksheet
    Dim dataArr As Variant
    Dim shtName As String

    Set allSht = Worksheets.Add(After:=Worksheets(Worksheets.Count))

    For isht = 1 To Worksheets.Count - 1
        With Worksheets(isht)
            dataArr = Intersect(.UsedRange, .Range("A:B")).Value
            shtName = .Name
        End With
        With allSht
            With .Cells(.Rows.Count, 1).End(xlUp).Offset(1)
                .Resize(UBound(dataArr, 1), UBound(dataArr, 2)).Value = dataArr
                .Offset(, 2) = "sheet_name"
                .Offset(1, 2).Resize(UBound(dataArr, 1) - 1).Value = shtName
            End With
        End With
    Next isht

    With allSht
        .Rows(1).Delete
        .Name = "All"
    End With
End Sub