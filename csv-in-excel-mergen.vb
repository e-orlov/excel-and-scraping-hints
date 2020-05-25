Sub ImportCSVFromFolder()
    Dim wsTemp As Worksheet, wsTarget As Worksheet, curCell As Range, CSVPFAD As String, fso As Object, f As Object, strCSVDelimiter As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "C:\Users\Evgeniy.MEDIAWORXDE\Downloads\1"
        .Title = "Ordnerauswahl"
        .ButtonName = "Auswahl..."
        .InitialView = msoFileDialogViewList
        If .Show = -1 Then
            CSVPFAD = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    'Legt das CSV-Trennzeichen für die Dateien fest
    strCSVDelimiter = ";"
    
    Set fso = CreateObject("Scripting.Filesystemobject")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Zielarbeitsblatt für die importierten Daten
    Set wsTarget = Worksheets(1)
    wsTarget.Name = "Zusammenfassung"
    'temporäres Arbeitsblatt für den Import der Daten erstellen
    Set wsTemp = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    
    'Inhalt des Zusammenfassungsblattes löschen
    wsTarget.UsedRange.Clear
    
    'Startausgabezelle festlegen
    Set curCell = wsTarget.Range("A1")
    For Each f In fso.GetFolder(CSVPFAD).Files
        If LCase(fso.GetExtensionName(f.Name)) = "csv" Then
            'Temporäres Sheet löschen
            wsTemp.UsedRange.Clear
            'CSV-Daten in Temporäres Sheet importieren
            With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & f.Path, Destination:=wsTemp.Range("$A$1"))
                .Name = "import"
                .FieldNames = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .TextFilePlatform = xlWindows
                .TextFileStartRow = 1
                .TextFileParseType = xlDelimited
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileOtherDelimiter = strCSVDelimiter
                .Refresh BackgroundQuery:=False
                .Delete
            End With
            
            With wsTemp
                'Daten in Zielsheet kopieren
                .UsedRange.Copy curCell
            End With
            'Ausgabezeile eins nach unten schieben
            Set curCell = wsTarget.Cells(wsTarget.UsedRange.Rows.Count + 2, 1)
        End If
    Next
    'Temporäres Sheet löschen
    wsTemp.Delete
    'Spalten anpassen
    wsTarget.Columns.AutoFit
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Vorgang beendet!", vbInformation
    Set fso = Nothing
End Sub
