Sub simpleXlsMerger()
Dim bookList As Workbook
Dim mergeObj As Object, dirObj As Object, filesObj As Object, everyObj As Object
Application.ScreenUpdating = False
Set mergeObj = CreateObject("Scripting.FileSystemObject")
 
'change folder path of excel files here
Set dirObj = mergeObj.Getfolder("U:\iMacros\Downloads")
Set filesObj = dirObj.Files
For Each everyObj In filesObj
Set bookList = Workbooks.Open(everyObj)
 
'change "A2" with cell reference of start point for every files here
'for example "B3:IV" to merge all files start from columns B and rows 3
'If you're files using more than IV column, change it to the latest column
'Also change "A" column on "A65536" to the same column as start point
Range("A2:IV" & Range("A65536").End(xlUp).Row).Copy
ThisWorkbook.Worksheets(1).Activate
 
'Do not change the following column. It's not the same column as above
Range("A65536").End(xlUp).Offset(1, 0).PasteSpecial
Application.CutCopyMode = False
bookList.Close
Next
End Sub