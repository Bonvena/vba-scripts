Option Explicit
Sub CTEA_1P1_Calculation()

'Calculates the overall and program-by-program 1p1 metric for the CLNA.
Dim LastRow, LastColumn, LastRow1p1, LastColumn1p1 As Long
Dim i, sp, numerator, numerator1, numerator2, numerator_intersection, denominator, denominator1, denominator2, num_total, denom_total, num_NTG_total, denom_NTG_total, enrollment, enrollment_total, completion, completion_total As Long
Dim rw, index, response, genderColumn, yrend_statColumn, programColumn, cipColumn, studentColumn, EMPSTAT_Column, EDUCSTAT_Column, nonTradMale_Column, nonTradFemale_Column As Long
Dim welcome, timeColumn, concentratorMin, firstSP, lastSP, ctea_type, completer, completer_notenrolled, other_notenrolled As Long
Dim lookupArray, empstatRange As Range
Dim result, target_1p1, overall_total, overall_NTG_total As Double
Dim uniqueProgramArray, empstatArray As Variant
Dim FirstSheet, Sheet1p1, sFile, nontrad_male, nontrad_female As String
Dim inputQuestionCLNA_Type, inputQuestion_1p1, fileTypes As String
Dim original_wb, crosswalk_wb As Workbook
Dim original_sht, crosswalk_sht, target1p1_ws As Worksheet

Set original_wb = ActiveWorkbook
FirstSheet = ActiveSheet.Name
Set original_sht = ActiveSheet

'asks users to select nontraditional crosswalk
welcome = MsgBox(prompt:="This script will provide 1P1 metric per program and special population for the CLNA." _
            & vbCrLf & vbCrLf & "It is intended to work on the CTEA files formatted for the IDEx system." _
            & vbCrLf & "The CTEA A file (1 OR 2) should be the only Excel workbook open." _
            & vbCrLf & "You must also run this as a Module (Insert -> Module in the VBA Editor)." _
            & vbCrLf & vbCrLf & "You should also have saved the 2020 Nontraditional Crosswalk where it can be located as well as the CTEA B file. " _
            & " As this file will be modified, you may wish to save a copy first before continuing. Otherwise, click OK." _
            , Title:="CTEA CLNA Calculation", Buttons:=vbOKCancel)

If welcome = vbCancel Then
    Exit Sub
End If

'asks if data is for CTEA 1 or 2
inputQuestionCLNA_Type = "Enter 1 if data file is for CTEA 1 (credit bearing courses) or 2 for" _
                        & " CTEA 2 (non-credit bearing courses) (Default is 1)."
                        
    'loop for incorrect number
    Do
    
        ctea_type = Application.InputBox(inputQuestionCLNA_Type, "CTEA Type", 1, , , , , 1)
        
        If ctea_type = False Then
            Exit Sub
        End If
    
    Loop While (ctea_type >= 3) Or (ctea_type <= 0)

If ctea_type = 1 Then

    studentColumn = 2 'Column of Student ID
    programColumn = 3 'Column of IRP Code
    cipColumn = 4 'Column of CIP Code
    timeColumn = 17 'Column of Credits Earned
    concentratorMin = 12 'Min number of credits to be considered a concentrator
    firstSP = 8 'Column of first special population to be analyzed
    lastSP = 16 ' Column of last special population (assumes all special populations are between firstSP and lastSP)
    genderColumn = 5 'Column of Gender
    yrend_statColumn = 18 ' Column of YRENDSTAT_ID
    completer_notenrolled = 6 'Received certificate/degree -- not enrolled during academic year
    completer = 4 'Completer, enrolled during academic year
    other_notenrolled = 5 'Other/Not Enrolled
    nonTradFemale_Column = 22 'Column that indicates whether the non-traditional gender of program is female.
    nonTradMale_Column = 23 'Column that indicates whether the non-traditional gender of program is male.
    EMPSTAT_Column = 24 'Column of employement status variable added from B file.
    EDUCSTAT_Column = 25 'Column of further education status variable added from B file.

Else
    
    studentColumn = 2 'Column of Student ID
    programColumn = 4 'Column of Program Code
    cipColumn = 5 'Column of CIP Code
    timeColumn = 18 'Column of Contact Hours
    concentratorMin = 100 'Min number of contact hours to be considered a concentrator
    firstSP = 9 'Column of first special population to be analyzed
    lastSP = 17 ' Column of last special population (assumes all special populations are between firstSP and lastSP)
    genderColumn = 6 'Column of Gender
    yrend_statColumn = 19 ' Column of YRENDSTAT_ID
    completer_notenrolled = 1 'Received certificate/degree -- not enrolled during academic year
    completer = 3 'Completer, enrolled during academic year
    other_notenrolled = 4 'Other/Not Enrolled
    nonTradFemale_Column = 22 'Column that indicates whether the non-traditional gender of program is female.
    nonTradMale_Column = 23 'Column that indicates whether the non-traditional gender of program is male.
    EMPSTAT_Column = 24 'Column of employement status variable added from B file.
    EDUCSTAT_Column = 25 'Column of further education status variable added from B file.

End If

'asks for Target 1p1 Rate
inputQuestion_1p1 = "Enter Target 1p1 Rate in number form without % (e.g. 50.45, not 50.45%)."
       
    'loops for incorrect percentage
    Do
    
        target_1p1 = Application.InputBox(inputQuestion_1p1, "Target 1p1 Rate", 50#, , , , , 1)
    
    Loop While (target_1p1 >= 100) Or (target_1p1 < 0)

If target_1p1 = False Then
    Exit Sub
Else
    target_1p1 = target_1p1 / 100
End If


'finds last row and column
LastRow = Worksheets(FirstSheet).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LastColumn = Worksheets(FirstSheet).Cells(1, Columns.Count).End(xlToLeft).Column

'changes format of all cells from the Gender field onward to General
Range(Cells(1, genderColumn), Cells(LastRow, LastColumn)).Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

Range("A1").Select
Application.CutCopyMode = False

'asks users to select nontraditional crosswalk
response = MsgBox(prompt:="Press OK to select the file containing the nontraditional crosswalk." & _
            vbCrLf & "CIP Codes should be in first column, nontraditional statuses in columns 4 and 5." _
            & vbCrLf & "If you do not have this, press Cancel to Exit.", Title:="CTEA CLNA Calculation", Buttons:=vbOKCancel)

If response = vbCancel Then
    Exit Sub
End If

fileTypes = "Excel Files (*.xls*) , *.xls*," & _
            "Text Files (*.txt; *.csv) , *.txt"

sFile = Application.GetOpenFilename(FileFilter:=fileTypes, Title:="Select nontraditional crosswalk file")

Application.ScreenUpdating = False

'lookups nontraditional female and male info and returns it to original file as new columns
Set crosswalk_wb = Workbooks.Open(sFile)
Set lookupArray = crosswalk_wb.Worksheets(1).Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 5))

With original_wb.Worksheets(FirstSheet)

    For rw = 2 To LastRow
        .Cells(rw, LastColumn + 1) = Trim(WorksheetFunction.IfError(Application.VLookup(.Cells(rw, cipColumn).Value2 & "", lookupArray, 4, False), "N/A"))
        .Cells(rw, LastColumn + 2) = Trim(WorksheetFunction.IfError(Application.VLookup(.Cells(rw, cipColumn).Value2 & "", lookupArray, 5, False), "N/A"))
    Next rw

End With

crosswalk_wb.Close savechanges:=False

Worksheets(FirstSheet).Activate

Cells(1, LastColumn + 1).Value = "Nontraditional for females"
Columns(LastColumn + 1).ColumnWidth = 30
Cells(1, LastColumn + 2).Value = "Nontraditional for males"
Columns(LastColumn + 2).ColumnWidth = 30

Worksheets(FirstSheet).Range(Worksheets(FirstSheet).Cells(2, LastColumn + 1), Worksheets(FirstSheet).Cells(LastRow, LastColumn + 2)).Replace 0, "N"

'finished with nontraditional lookup
'creates a worksheet called "Ignore" and adds it to the end
Dim Ws As Worksheet
Set Ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))

Ws.Name = "Ignore"

'uses AutoFilter to take unique IRP codes and copy to Ignore sheet. Sorts and then adds to array
Worksheets(FirstSheet).Range(Worksheets(FirstSheet).Cells(1, programColumn), Worksheets(FirstSheet).Cells(LastRow, programColumn)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Ignore").Range("A1"), Unique:=True
LastRow = Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row

'clear the sorted field and apply AutoFilter
Ws.Range("A1:A" & LastRow).Select
Ws.Range("A1:A" & LastRow).AutoFilter
Ws.AutoFilter.Sort.SortFields.Clear

Ws.AutoFilter.Sort.SortFields.Add Order:=xlAscending, _
    SortOn:=xlSortOnValues, Key:=Range("A1:A" & LastRow)

Ws.AutoFilter.Sort.Apply

'create array of unique IRP codes
uniqueProgramArray = Application.Transpose(Ws.Range("A2:A" & LastRow))

'turn off AutoFilter
ActiveSheet.AutoFilterMode = False

'delete Ignore Worksheet
Application.DisplayAlerts = False
Worksheets("Ignore").Delete
Application.DisplayAlerts = True

LastRow = Worksheets(FirstSheet).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LastColumn = Worksheets(FirstSheet).Cells(1, Columns.Count).End(xlToLeft).Column

'looks up EMPSTAT_ID and EDUCSTAT_ID from B file
response = MsgBox(prompt:="Press OK to select the file containing the B file corresponding to the above A file." & _
            vbCrLf & "If you do not have this, press Cancel to Exit.", Title:="CTEA B File", Buttons:=vbOKCancel)

If response = vbCancel Then
    Exit Sub
End If

fileTypes = "Excel Files (*.xls*) , *.xls*," & _
            "Text Files (*.txt; *.csv) , *.txt"

sFile = Application.GetOpenFilename(FileFilter:=fileTypes, Title:="Select corresponding B file for lookup")


'lookups EMPSTAT and EDUCSTAT info and returns it to original file as new columns
Set crosswalk_wb = Workbooks.Open(sFile)
Set lookupArray = crosswalk_wb.Worksheets(1).Range(Cells(1, 2), Cells(Cells(Rows.Count, 2).End(xlUp).Row, 4))


With original_wb.Worksheets(FirstSheet)

    For rw = 2 To LastRow
        .Cells(rw, LastColumn + 1) = Trim(WorksheetFunction.IfError(Application.VLookup(.Cells(rw, studentColumn).Value2 & "", lookupArray, 2, False), "N/A"))
        .Cells(rw, LastColumn + 2) = Trim(WorksheetFunction.IfError(Application.VLookup(.Cells(rw, studentColumn).Value2 & "", lookupArray, 3, False), "N/A"))
    Next rw

End With

crosswalk_wb.Close savechanges:=False

Worksheets(FirstSheet).Activate

Cells(1, LastColumn + 1).Value = "EMPSTAT_ID"
Columns(LastColumn + 1).ColumnWidth = 20
Cells(1, LastColumn + 2).Value = "EDUCSTAT_ID"
Columns(LastColumn + 2).ColumnWidth = 30

'***1P1 TABLE START***
'Create table for 1p1 and adjust column widths
Set target1p1_ws = original_wb.Worksheets.Add(Type:=xlWorksheet, After:=Application.ActiveSheet)
target1p1_ws.Name = "1p1"
Sheet1p1 = target1p1_ws.Name

With Worksheets(Sheet1p1)

    .Cells(1, 1).Value = "CIP Code"
    .Columns(1).ColumnWidth = 15
    .Cells(1, 2).Value = "Program/IRP Code"
    .Columns(2).ColumnWidth = 20
    .Cells(1, 3).Value = "Overall"
    .Columns(3).ColumnWidth = 40
    
End With

LastColumn1p1 = target1p1_ws.Cells(1, Columns.Count).End(xlToLeft).Column

'adds headers to Sheet1p1
i = 1
For sp = firstSP To lastSP
    Worksheets(Sheet1p1).Cells(1, LastColumn1p1 + i).Value2 = Worksheets(FirstSheet).Cells(1, sp).Value2
    Worksheets(Sheet1p1).Columns(LastColumn1p1 + i).ColumnWidth = 5 + Len(Worksheets(FirstSheet).Cells(1, sp).Value2)
    i = i + 1
Next sp


'calculation of overall 1p1 for each program VERIFIED WORKING CORRECTLY
num_total = 0
denom_total = 0
For i = LBound(uniqueProgramArray) To UBound(uniqueProgramArray)
    
    Worksheets(Sheet1p1).Cells(i + 1, 2).Value2 = uniqueProgramArray(i)
    
    With Worksheets(FirstSheet)
       
       'numerator consists of ALL COMPLETERS in the program (during and not during AY) employed or pursuing further education
        'numerator1 looks if completer was pursuing further education
        numerator1 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1)) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1))
                    
                            
        'numerator2 looks if completer was employed
        numerator2 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10)))
        
                
        'numerator_intersection is the intersection of the two sets (PURSUING EDUC AND EMPLOYED)
        numerator_intersection = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10)))
        
               
        'combined both numerators as OR condition for union of sets
        numerator = numerator1 + numerator2 - numerator_intersection
        
        'denominator consists of ALL COMPLETERS in the program
        denominator = WorksheetFunction.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i)) + _
                    WorksheetFunction.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i))
                    

    End With
    
    'adds this to global variables for overall total calculation later on
    num_total = num_total + numerator
    denom_total = denom_total + denominator
    
    'checks divide by zero error and formats result
    If denominator = 0 Then
        Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1).Value = "N/A (no cases)"
    Else
        result = numerator / denominator

        If result >= target_1p1 Then
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1).Value = FormatPercent(result, 2) & " (" & numerator & "/" & denominator & ")" & " (Above Target of " & Format(target_1p1, "0.00%") & ")"
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1).Font.Color = RGB(0, 110, 0)
        Else
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1).Value = FormatPercent(result, 2) & " (" & numerator & "/" & denominator & ")" & " (Below Target of " & Format(target_1p1, "0.00%") & ")"
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1).Font.Color = RGB(230, 0, 0)
        End If

    End If

Next

'calculation of 1p1 for each program's special population VERIFIED WORKING CORRECTLY
For i = LBound(uniqueProgramArray) To UBound(uniqueProgramArray)
    
    index = 1
    For sp = firstSP To lastSP

        With Worksheets(FirstSheet)
            
            
            'numerator consists of ALL COMPLETERS in the program IN THE S.P. (during and not during AY) employed OR pursuing further education
            'numerator1 looks if completer IN THE S.P. was pursuing further education
            numerator1 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1)) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1))
            
                        
            'numerator2 looks if completer IN THE S.P. was employed
            numerator2 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10)))
    
                        
            'numerator_intersection is the intersection of the two sets (PURSUING EDUC AND EMPLOYED)
            numerator_intersection = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10)))
        
            
                    
            'combined both numerators as OR condition for union of sets
            numerator = numerator1 + numerator2 - numerator_intersection
            
            
            'denominator consists of ALL COMPLETERS in the program IN THE S.P.
            denominator = WorksheetFunction.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                        .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                        .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1) + _
                        WorksheetFunction.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                        .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                        .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1)
    
        End With
        
        'checks divide by zero error and formats result
        If denominator = 0 Then
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + index).Value = "N/A (no cases)"
        Else
            result = numerator / denominator
    
            If result >= target_1p1 Then
                Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + index).Value = FormatPercent(result, 2) & " (" & numerator & "/" & denominator & ")" & " (Above Target of " & Format(target_1p1, "0.00%") & ")"
                Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + index).Font.Color = RGB(0, 110, 0)
            Else
                Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + index).Value = FormatPercent(result, 2) & " (" & numerator & "/" & denominator & ")" & " (Below Target of " & Format(target_1p1, "0.00%") & ")"
                Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + index).Font.Color = RGB(230, 0, 0)
            End If
    
        End If
    
        index = index + 1
        
    Next sp

Next i


'adds CIP code to new table
For i = LBound(uniqueProgramArray) To UBound(uniqueProgramArray)
    
    With Worksheets(FirstSheet)
    
        Worksheets(Sheet1p1).Cells(i + 1, 1).Value2 = WorksheetFunction.index(.Range(.Cells(2, cipColumn), .Cells(LastRow, cipColumn)), _
        WorksheetFunction.Match(uniqueProgramArray(i), .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), 0))
    
    End With

Next i


LastColumn1p1 = target1p1_ws.Cells(1, Columns.Count).End(xlToLeft).Column
LastRow1p1 = target1p1_ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row


'calculates 1p1 rate for each program for nontraditional gender, which is not a S.P. column in the file.
num_NTG_total = 0
denom_NTG_total = 0
For i = LBound(uniqueProgramArray) To UBound(uniqueProgramArray)
    
    With Worksheets(FirstSheet)
        
        'looks up the nontraditional gender for the program
        nontrad_female = WorksheetFunction.IfError(WorksheetFunction.index(.Range(.Cells(2, nonTradFemale_Column), .Cells(LastRow, nonTradFemale_Column)), _
                        WorksheetFunction.Match(uniqueProgramArray(i), .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), 0)), "")
        
        nontrad_male = WorksheetFunction.IfError(WorksheetFunction.index(.Range(.Cells(2, nonTradMale_Column), .Cells(LastRow, nonTradMale_Column)), _
                        WorksheetFunction.Match(uniqueProgramArray(i), .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), 0)), "")
        
                
        'if the nontrad. gender is female, counts only female and non-binary and unknown COMPLETERS
        If nontrad_female = "Y" Then
        
            'numerator consists of ALL COMPLETERS in the program IN THE NON-TRAD GENDERS (during and not during AY) employed OR pursuing further education
            'numerator1 looks if completer IN THE NON-TRAD GENDERS was pursuing further education
            numerator1 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1)) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1))
            
                        
        
            'numerator2 looks if completer IN THE NON-TRAD GENDERS was employed
            numerator2 = Application.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4), _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10)))) + _
                    Application.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4), _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10))))
        
                       
            
            'numerator_intersection is the intersection of the two sets (PURSUING EDUC AND EMPLOYED)
            numerator_intersection = Application.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10)))) + _
                    Application.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10))))
            
            'combined both numerators as OR condition for union of sets
            numerator = numerator1 + numerator2 - numerator_intersection
            
            
            'denominator consists of ALL female/non-binary/unknown in the program COMPLETERS
            denominator = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(2, 3, 4)))
        
            'adds this to global variables for overall total nontrad. calculation later on
            num_NTG_total = num_NTG_total + numerator
            denom_NTG_total = denom_NTG_total + denominator
        
        'if the nontrad. gender is male, counts only male and non-binary and unknown COMPLETERS
        ElseIf nontrad_male = "Y" Then
            
            'numerator consists of ALL COMPLETERS in the program IN THE NON-TRAD GENDERS (during and not during AY) employed OR pursuing further education
            'numerator1 looks if completer IN THE NON-TRAD GENDERS was pursuing further education
            numerator1 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1)) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1))
        
                   
            'numerator2 looks if completer IN THE NON-TRAD GENDERS was employed
            numerator2 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4), _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10)))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4), _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10))))
        
                        
            'numerator_intersection is the intersection of the two sets (PURSUING EDUC AND EMPLOYED)
            numerator_intersection = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10)))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4), _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Application.Transpose(Array(1, 2, 3, 7, 8, 9, 10))))
            
            'combined both numerators as OR condition for union of sets
            numerator = numerator1 + numerator2 - numerator_intersection
            
                        
            'denominator consists of male/non-binary/unknown in the program COMPLETERS
            denominator = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, programColumn), .Cells(LastRow, programColumn)), uniqueProgramArray(i), _
                    .Range(.Cells(2, genderColumn), .Cells(LastRow, genderColumn)), Array(1, 3, 4)))
            
            'adds this to global variables for overall total nontrad. calculation later on
            num_NTG_total = num_NTG_total + numerator
            denom_NTG_total = denom_NTG_total + denominator
            
        'otherwise there is no nontraditional gender for program based on CIP code
        Else: denominator = -1
        
        End If
            
    End With
    
    'checks for divide by zero error and no nontraditional gender. Otherwise formats result
    If denominator = 0 Then
        Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + 1).Value = "N/A (no cases)"
    ElseIf denominator = -1 Then
        Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + 1).Value = "N/A (no nontraditional gender)"
    Else
        result = numerator / denominator

        If result >= target_1p1 Then
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + 1).Value = FormatPercent(result, 2) & " (" & numerator & "/" & Format(denominator, "#,###") & ")" & " (Above Target of " & Format(target_1p1, "0.00%") & ")"
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + 1).Font.Color = RGB(0, 110, 0)
        Else
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + 1).Value = FormatPercent(result, 2) & " (" & numerator & "/" & Format(denominator, "#,###") & ")" & " (Below Target of " & Format(target_1p1, "0.00%") & ")"
            Worksheets(Sheet1p1).Cells(i + 1, LastColumn1p1 + 1).Font.Color = RGB(230, 0, 0)
        End If

    End If
    
        
Next i

Worksheets(Sheet1p1).Cells(1, LastColumn1p1 + 1).Value = "Nontraditional in this field"
Worksheets(Sheet1p1).Cells(1, LastColumn1p1 + 1).ColumnWidth = 75

LastColumn1p1 = target1p1_ws.Cells(1, Columns.Count).End(xlToLeft).Column


'calculation of overall 1p1 for college (does not remove duplicates)
'checks divide by zero error and formats result
If denom_total = 0 Then
    Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3).Value = "N/A (no cases)"
Else
    overall_total = num_total / denom_total

    If overall_total >= target_1p1 Then
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3).Value = FormatPercent(overall_total, 2) & " (" & num_total & "/" & Format(denom_total, "#,###") & ")" & " (Above Target of " & Format(target_1p1, "0.00%") & ")"
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3).Font.Color = RGB(0, 110, 0)
    Else
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3).Value = FormatPercent(overall_total, 2) & " (" & num_total & "/" & Format(denom_total, "#,###") & ")" & " (Below Target of " & Format(target_1p1, "0.00%") & ")"
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3).Font.Color = RGB(230, 0, 0)
    End If

End If

Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 1).Value = "Overall (does not check for duplicates)"
Worksheets(Sheet1p1).Range(Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 1), Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 2)).Merge Across:=True


'calculation of overall 1p1 for each special population (does not remove duplicates)
index = 1
For sp = firstSP To lastSP

    With Worksheets(FirstSheet)
        
        'numerator consists of ALL COMPLETERS IN THE S.P. (during and not during AY) employed OR pursuing further education
        'numerator1 looks if completer IN THE S.P. was pursuing further education
        numerator1 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1)) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1))
        
        
        'numerator2 looks if completer IN THE S.P. was employed
        numerator2 = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, _
                    .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10)))
        
        
        'numerator_intersection is the intersection of the two sets (PURSUING EDUC AND EMPLOYED IN THE S.P.)
        numerator_intersection = WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10))) + _
                    WorksheetFunction.Sum(Application.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                    .Range(.Cells(2, EDUCSTAT_Column), .Cells(LastRow, EDUCSTAT_Column)), 1, _
                    .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1, .Range(.Cells(2, EMPSTAT_Column), .Cells(LastRow, EMPSTAT_Column)), Array(1, 2, 3, 7, 8, 9, 10)))
        
        'combined both numerators as OR condition for union of sets
        numerator = numerator1 + numerator2 - numerator_intersection
        
                   
        'denominator consists of ALL COMPLETERS IN THE S.P.
        denominator = WorksheetFunction.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer_notenrolled, _
                        .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1) + _
                        WorksheetFunction.CountIfs(.Range(.Cells(2, yrend_statColumn), .Cells(LastRow, yrend_statColumn)), completer, _
                        .Range(.Cells(2, sp), .Cells(LastRow, sp)), 1)
    
    
    End With
    
    'checks divide by zero error and formats result
    If denominator = 0 Then
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3 + index).Value = "N/A (no cases)"
    Else
        result = numerator / denominator

        If result >= target_1p1 Then
            Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3 + index).Value = FormatPercent(result, 2) & " (" & numerator & "/" & Format(denominator, "#,###") & ")" & " (Above Target of " & Format(target_1p1, "0.00%") & ")"
            Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3 + index).Font.Color = RGB(0, 110, 0)
        Else
            Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3 + index).Value = FormatPercent(result, 2) & " (" & numerator & "/" & Format(denominator, "#,###") & ")" & " (Below Target of " & Format(target_1p1, "0.00%") & ")"
            Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, 3 + index).Font.Color = RGB(230, 0, 0)
        End If

    End If

    index = index + 1
    
Next sp


'calculation of overall nontraditional gender 1p1 for college (does not remove duplicates)
'checks divide by zero error and formats result
If denom_NTG_total = 0 Then
    Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, LastColumn1p1).Value = "N/A (no cases)"
Else
    overall_NTG_total = num_NTG_total / denom_NTG_total

    If overall_NTG_total >= target_1p1 Then
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, LastColumn1p1).Value = FormatPercent(overall_NTG_total, 2) & " (" & num_NTG_total & "/" & Format(denom_NTG_total, "#,###") & ")" & " (Above Target of " & Format(target_1p1, "0.00%") & ")" _
        & " (Excludes programs with no nontrad. gender)"
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, LastColumn1p1).Font.Color = RGB(0, 110, 0)
    Else
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, LastColumn1p1).Value = FormatPercent(overall_NTG_total, 2) & " (" & num_NTG_total & "/" & Format(denom_NTG_total, "#,###") & ")" & " (Below Target of " & Format(target_1p1, "0.00%") & ")" _
        & " (Excludes programs with no nontrad. gender)"
        Worksheets(Sheet1p1).Cells(LastRow1p1 + 1, LastColumn1p1).Font.Color = RGB(230, 0, 0)
    End If

End If

LastRow1p1 = target1p1_ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LastColumn1p1 = target1p1_ws.Cells(1, Columns.Count).End(xlToLeft).Column

'formats new table
With Worksheets(Sheet1p1)

    .Range(.Cells(1, 1), .Cells(LastRow1p1, LastColumn1p1)).HorizontalAlignment = xlLeft
    .Range(.Cells(1, 1), .Cells(LastRow1p1, LastColumn1p1)).Borders.LineStyle = xlContinuous
    .Range(.Cells(1, 1), .Cells(LastRow1p1, LastColumn1p1)).BorderAround Weight:=xlThick

End With

Worksheets(Sheet1p1).Activate
Worksheets(Sheet1p1).Range("A1").Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
