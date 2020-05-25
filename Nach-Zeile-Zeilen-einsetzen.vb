Sub AddRows()
    ScreenUpdating = False

    With ActiveSheet
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

    Dim AddRows As Integer: AddRows = 10

    Dim i As Integer: i = lastrow

    Do While i <> 1
        Rows(i & ":" & i + AddRows - 1).Insert
        i = i - 1
    Loop

    ScreenUpdating = True
End Sub