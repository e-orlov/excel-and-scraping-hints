Sub Start()
Dim ArData, ArFile(), ArAusgabe(), n&, nn&, nnn&, nCount&
Dim oDic As Object, oApp As Excel.Application
Dim sPath$, tmpFileName$

sPath = "C:\Users\eorlov\Downloads\as24-analytics" 'Pfad anpassen ********** 

sPath = IIf(Right$(sPath, 1) <> "\", sPath & "\", sPath)
tmpFileName = Dir(sPath & "*.xls?", vbNormal)
Do While tmpFileName <> ""
    Redim Preserve ArFile(n)
    ArFile(n) = sPath & tmpFileName
    n = n + 1
    tmpFileName = Dir()
Loop
If n < 1 Then Exit Sub 'keine Datei gefunden ************* 

Set oApp = New Excel.Application

Set oDic = CreateObject("Scripting.Dictionary")
With oApp
    .ScreenUpdating = False
    .EnableEvents = False
    .DisplayAlerts = False
    
    For n = Lbound(ArFile) To Ubound(ArFile)
        Application.StatusBar = "Lese Datei " & n + 1 & " von " & Ubound(ArFile) + 1
        With .Workbooks.Open(Filename:=ArFile(n), ReadOnly:=True)
            With .Sheets(1) 'evtl. anpassen 
                nn = .Cells(.Rows.Count, 1).End(xlUp).Row
                If nn > 1 Then
                    ArData = .Range("A2", .Cells(nn, 1)).Resize(, 19) 'bis Spalte S 
                End If
            End With
            .Close False
        End With
        If IsArray(ArData) Then
            For nn = 1 To Ubound(ArData)
                If Not oDic.exists(ArData(nn, 1)) Then
                    nCount = nCount + 1
                    Redim Preserve ArAusgabe(1 To 20, 1 To nCount)
                    For nnn = 2 To Ubound(ArData, 2)
                        ArAusgabe(nnn + 1, nCount) = ArData(nn, nnn)
                    Next nnn
                    ArAusgabe(1, nCount) = ArData(nn, 1)
                End If
                oDic(ArData(nn, 1)) = oDic(ArData(nn, 1)) + 1
            Next nn
            ArData = Empty
        End If
    Next n

    .ScreenUpdating = True
    .EnableEvents = True
    .DisplayAlerts = True
    .Quit
End With
Set oApp = Nothing
Application.StatusBar = False
If oDic.Count > 0 Then
    ArAusgabe = TransposeData(ArAusgabe, oDic)
    With ThisWorkbook.Sheets.Add  ' neue Tabelle erstellen ********************* 
        .Range("A2").Resize(Ubound(ArAusgabe), Ubound(ArAusgabe, 2)) = ArAusgabe
    End With
End If
MsgBox "fertig"
Set oDic = Nothing
End Sub

Function TransposeData(ArValues, oDic As Object)
Dim n&, nn&, NewAr()
Redim Preserve NewAr(1 To Ubound(ArValues, 2), 1 To Ubound(ArValues))
For n = Lbound(ArValues, 2) To Ubound(ArValues, 2)
    For nn = Lbound(ArValues) To Ubound(ArValues)
        NewAr(n, nn) = ArValues(nn, n)
    Next nn
    NewAr(n, 2) = oDic(NewAr(n, 1))
Next n
TransposeData = NewAr
End Function