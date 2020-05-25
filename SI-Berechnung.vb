Sub Makro1()
'
' Makro1 Makro
'

'
    Application.CutCopyMode = False
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\Evgeniy.MEDIAWORXDE\Downloads\Versicherungen-Research\Project_ Versicherungen SI (1)\domains\advigon.com.csv" _
        , Destination:=Range("$A$1"))

        .Name = "advigon.com"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Range("A5").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    ChDir "C:\Users\Evgeniy.MEDIAWORXDE\Downloads"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Evgeniy.MEDIAWORXDE\Downloads\advigon.com.xlsm", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    
End Sub
Sub Makro2()
'
' Makro2 Makro
'

'
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Brand"
    Range("G2").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(C[-6],'[Versicherungen-Brand-Non-Brand-KWs-SV.xlsx]Brand+Non-Brand Keywords mit SV'!C1:C3,3,0)"
End Sub
Sub Makro3()
'
' Makro3 Makro
'

'
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "SI"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-7],'[Versicherungen-Brand-Non-Brand-KWs-SV.xlsx]Brand+Non-Brand Keywords mit SV'!C1:C2,2,0))/RC[-2])+(IF(RC[-2]=1,""33,9"",IF(RC[-2]=2,""16,28"",IF(RC[-2]=3,""10,36"",IF(RC[-2]=4,""7"",IF(RC[-2]=5,""5,64"")))))+IF(RC[-2]=6,""4,13"",IF(RC[-2]=7,""3,27"",IF(RC[-2]=8,""2,61"",IF(RC[-2]=9,""2,18"",IF(RC[-2]=10,""1,82"")))))+IF(RC[-2]=11,""1,77"",I" & _
        "F(RC[-2]=12,""1,81"",IF(RC[-2]=13,""1,85"",IF(RC[-2]=14,""1,9"",IF(RC[-2]=15,""2,04"")))))+IF(RC[-2]=16,""1,68"",IF(RC[-2]=17,""1,61"",IF(RC[-2]=18,""1,65"",IF(RC[-2]=19,""1,62"",IF(RC[-2]=20,""1,59"",""0"")))))),0)" & _
        ""
    Range("H3").Select
End Sub
Sub Makro4()
'
' Makro4 Makro
'

'
    Range("J1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabelle1!R1C7:R1048576C8", Version:=6).CreatePivotTable TableDestination:= _
        "Tabelle1!R1C10", TableName:="PivotTable5", DefaultVersion:=6
    Sheets("Tabelle1").Select
    Cells(1, 10).Select
    With ActiveSheet.PivotTables("PivotTable5")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable5").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Brand")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("SI"), "Summe von SI", xlSum

End Sub
Sub Makro5()
'
' Makro5 Makro
'

'
    Range("J2:K6").Select
    Selection.Copy
    Windows("si.xlsx").Activate
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.Save

End Sub
