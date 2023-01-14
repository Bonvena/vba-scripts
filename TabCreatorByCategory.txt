Sub TabCreatorByCategory()

Application.ScreenUpdating = False

Dim colName, inputQuestionColName As String
Dim colNum, LastColumn, welcome As Long
Dim uniqueArray As Variant, LastRow As Long, val As Variant, SheetRange As Range, errMsg As String

'Welcome message for beginner
welcome = MsgBox(prompt:="This script will create separate tabs for each category of a variable column in a file." _
            & vbCrLf & vbCrLf & "It is not intended for scale variables like ID." _
            & vbCrLf & "The file in question should be the only Excel workbook open." _
            & " The script will work on the current tab of the file regardless of the number of other tabs. " _
            & " As this file will be modified, you may wish to save a copy first before continuing. Otherwise, click OK." _
            , Title:="Tab Creator by Category", Buttons:=vbOKCancel)

If welcome = vbCancel Then
    Exit Sub
End If

'asks for column of variable to filter on
inputQuestionColName = "Enter letter(s) of column to filter on without quotation marks (e.g. B, X, AC, DB, etc.)"

'loop for incorrect number
    Do
    
        colName = Application.InputBox(inputQuestionColName, "Column Name", , , , , , 2)
        
        If colName = False Then
            Exit Sub
        End If
        
    Loop While (Not colName Like "*[a-zA-Z]*")
   

colNum = Range(colName & 1).Column

'works on current sheet
Dim FirstSheet As Worksheet
Set FirstSheet = ActiveSheet

'creates a worksheet called "Ignore" and adds it to the end
Dim Ws As Worksheet
Set Ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))

Ws.Name = "Ignore"

'uses AutoFilter to take unique departments and copy to Ignore sheet. Sorts and then adds to array
FirstSheet.Columns(colNum).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Ignore").Range("A1"), Unique:=True
LastRow = Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row

'clear the sorted field and apply AutoFilter
Ws.Range("A1:A" & LastRow).Select
Ws.Range("A1:A" & LastRow).AutoFilter
Ws.AutoFilter.Sort.SortFields.Clear

Ws.AutoFilter.Sort.SortFields.Add Order:=xlAscending, _
    SortOn:=xlSortOnValues, Key:=Range("A1:A" & LastRow)

Ws.AutoFilter.Sort.Apply

'create array of unique departments
uniqueArray = Application.Transpose(Ws.Range("A2:A" & LastRow))

'resets LastRow and LastColumn to the sheet that user wants script to run on
LastRow = FirstSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastColumn = FirstSheet.Cells(1, Columns.Count).End(xlToLeft).Column

'select sheet to filter on
Set SheetRange = FirstSheet.Range(FirstSheet.Cells(1, 1), FirstSheet.Cells(LastRow, LastColumn))

'turn off AutoFilter
SheetRange.Parent.AutoFilterMode = False

'for each category in array, add new sheet with dataset belonging to that category
'quotes had to be added to sheet name because some names are reserved for Excel
For i = LBound(uniqueArray) To UBound(uniqueArray)

    SheetRange.AutoFilter Field:=colNum, Criteria1:=CStr(uniqueArray(i))
    Dim cat As Worksheet
    Set cat = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    cat.Name = Left(Chr(34) & CStr(uniqueArray(i)) & Chr(34), 30)

    SheetRange.Parent.AutoFilter.Range.Copy
    With cat.Range("A1")
        .PasteSpecial Paste:=8
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End With
Next

SheetRange.Parent.AutoFilterMode = False

'deletes helper sheet without displaying message
Application.DisplayAlerts = False
Worksheets("Ignore").Delete
Application.DisplayAlerts = True

Application.ScreenUpdating = True

End Sub
