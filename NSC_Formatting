' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' // NSC StudentTracker Formatting Macro
' //
' // For formatting files for National Student Clearinghouse's StudentTracker service.
' // This script will format CO, DA, or SE data queries for upload to National Student Clearinghouse's StudentTracker service.
' // To avoid repetitive dialogue windows, you will have to adjust the first few lines of code to your particular institution.
' // Original file must have first name, middle name, last name, YYYYMMDD birth date and student ID (Requester Return Field) in Columns A to E IN THAT ORDER,
' // with institution-specific variable names/headings in the first row.
' // (The code assumes the institution saves suffixes to the end of the last name field, not as its own separate field.
' // If not, the suffix column would have to be merged afterwards).
' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Sub NSC_FORMATTING()

Dim aFill As Long
Dim searchDate As Long
Dim creationDate As Long
Dim welcome As VbMsgBoxResult
Dim inputQuestion As String
Dim todayDate As String
Dim schoolName As String
Dim schoolCode As String
Dim branchCode As String
Dim queryOption As String
Dim rng As Range
Dim cell As Range
Dim numberCount As Integer
Dim wrongDateCount As Integer

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'CHANGE THESE LINES TO MATCH YOUR INSTITUTION.
schoolName = "ABC UNIVERSITY"
schoolCode = "000000"
branchCode = "00"
queryOption = "SE"

'Welcome message
welcome = MsgBox(prompt:="This script will format CO, DA, or SE queries for upload to National Student Clearinghouse's StudentTracker service." _
        & vbCrLf & vbCrLf & "Original file must have first name, middle name, last name, YYYYMMDD birth date and student ID in Columns A to E IN THAT ORDER, with headings." _
        & vbCrLf & vbCrLf & "It assumes you do not have a separate column for surname suffix." _
        & vbCrLf & "While it will format the file it will not remove errors, of which some of the most common will be listed on the right." _
        & " You must delete these extra columns once finished for a successful upload to NSC." _
        & vbCrLf & vbCrLf & "To avoid typing your institution's info each time they are hard coded into the beginning of the script." _
        & " You should adjust them in the Visual Basic Editor." _
        & vbCrLf & vbCrLf & "As this file will be modified, you may wish to save a copy first before continuing. Otherwise, click OK." _
        , Title:="NSC StudentTracker Formatting Macro", Buttons:=vbOKCancel)

If welcome = vbCancel Then
    Exit Sub
End If

Dim numberArray() As Variant
ReDim numberArray(1)
numberArray(1) = "Locations:"

Dim wrongDateArray() As Variant
ReDim wrongDateArray(1)
wrongDateArray(1) = "Locations:"

todayDate = Format(Date, "YYYYMMDD")
creationDate = todayDate
    
inputQuestion = "Enter NSC search start date as YYYYMMDD. Date cannot be in the future."
   
'loops for incorrect number, but does not check for everything or most things

Do
    searchDate = Application.InputBox(prompt:=inputQuestion, Title:="Search Start Date", Default:=Format(Date, "YYYYMMDD"), Type:=1)

Loop While searchDate > CLng(todayDate)
    
Columns("A:A").Select
Selection.Insert Shift:=xlToRight
Range("A2").Select
ActiveCell.FormulaR1C1 = "D1"
aFill = Range("B" & Rows.Count).End(xlUp).Row
Range("A2:A" & aFill).FillDown
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Columns("D:D").Select
Selection.Insert Shift:=xlToRight
Range("D2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1], 20)"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & aFill)
Range("D2:D" & aFill).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("C:C").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Columns("E:E").Select
Selection.Insert Shift:=xlToRight
Range("E2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1], 1)"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & aFill)
Range("E2:E" & aFill).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("D:D").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Columns("F:F").Select
Selection.Insert Shift:=xlToRight
Range("F2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1], 20)"
Range("F2").Select
Selection.AutoFill Destination:=Range("F2:F" & aFill)
Range("F2:F" & aFill).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("E:E").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
Columns("F:F").Select
Selection.Insert Shift:=xlToRight
Columns("H:H").Select
Selection.Insert Shift:=xlToRight
Range("H2").Select
Range("H2").Value = searchDate
Range("H2:H" & aFill).FillDown
Columns("I:I").Select
Selection.Insert Shift:=xlToRight
Cells.Select
Range("I1").Activate
Selection.NumberFormat = "@"
Columns("J:J").Select
Selection.Insert Shift:=xlToRight
Range("J2").Select
Range("J2").Value = schoolCode
Range("J2:J" & aFill).FillDown

Columns("K:K").Select
Selection.Insert Shift:=xlToRight
Range("K2").Select
Range("K2").Value = "00"
Range("K2:K" & aFill).FillDown
Range("K2:K" & aFill).Select
Range("L1").Select
Selection.ClearContents
Rows("1:1").Select
Selection.ClearContents
Range("A1").Select
ActiveCell.FormulaR1C1 = "H1"
Range("B1").Select
ActiveCell.FormulaR1C1 = schoolCode
Range("C1").Select
ActiveCell.FormulaR1C1 = branchCode
Range("D1").Select
ActiveCell.FormulaR1C1 = schoolName
Range("E1").Select
ActiveCell.FormulaR1C1 = creationDate
Range("F1").Select
ActiveCell.FormulaR1C1 = queryOption
Range("G1").Select
ActiveCell.FormulaR1C1 = "I"
Range("A1").Select
Selection.End(xlDown).Select

'Footer row
Range("A" & (aFill + 1)).Select
ActiveCell.FormulaR1C1 = "T1"
ActiveCell.Offset(0, 1).Select
Selection.NumberFormat = "General"
ActiveCell.FormulaR1C1 = "=ROW()"
Cells.Select
ActiveCell.Activate
Selection.NumberFormat = "@"
Range("A1").Select

'Adjusts column widths
Columns("A").ColumnWidth = 3
Columns("F:H").ColumnWidth = 9
Columns("I").ColumnWidth = 2
Columns("J").ColumnWidth = 7
Columns("K").ColumnWidth = 2
Columns("M").ColumnWidth = 7
Columns("N").ColumnWidth = 22

'Checks for special phrases that often result in errors when submitting to NSC
'Suffixes are often found after last name when they belong in separate column for NSC
'Periods after "St." are usually fine but others should be removed.
Range("N2") = "Count of suffixes and phrases that can result in NSC errors--correct manually"
Range("N3") = "JR"
Range("N4") = "SR"
Range("N5") = "II"
Range("N6") = "III"
Range("N7") = "IV"
Range("N8") = "NLN"
Range("N9") = "NFN"
Range("N10") = "Birth Year < 1910"
Range("N11") = "Periods (.)"
Range("N12") = "Underscores (_)"
Range("N13") = "Open Parenthesis ("
Range("N14") = "Close Parenthesis )"
Range("N15") = Chr(34) & Chr(45) & Chr(34) & " in Middle Name Field"
Range("N16") = "!"
Range("N17") = "?"
Range("N18") = "Any Numbers Saved as Text"

Range("O3") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "* JR*")
Range("O4") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "* SR*")
Range("O5") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "* II")
Range("O6") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "* III")
Range("O7") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "* IV")
Range("O8") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*NLN*")
Range("O9") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*NFN*")
Range("O11") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*.*")
Range("O12") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*_*")
Range("O13") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*(*")
Range("O14") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*)*")
Range("O15") = Application.WorksheetFunction.CountIf(Range("D2:D" & aFill), "*-*")
Range("O16") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*!*")
Range("O17") = Application.WorksheetFunction.CountIf(Range("C2:E" & aFill), "*~?*")

'Check to see if there are any birth dates that are not humanely possible for current students.
wrongDateCount = 0
Set rng = Range("G2:G" & aFill)
For Each cell In rng.Cells
    If cell.Value < 19100101 Then
        wrongDateCount = wrongDateCount + 1
        ReDim Preserve wrongDateArray(UBound(wrongDateArray) + 1)
        wrongDateArray(UBound(wrongDateArray)) = cell.Address
    End If
Next

Range("O10") = wrongDateCount

If wrongDateCount > 0 Then
    Dim Destination As Range
    Set Destination = Range("P10")
    Set Destination = Destination.Resize(1, UBound(wrongDateArray) + 1)
    Destination.Value = wrongDateArray
End If

'Check to see if there are any values in the name columns that are numbers, an error.
numberCount = 0
Set rng = Range("C2:E" & aFill)
For Each cell In rng.Cells
    If IsNumeric(cell) Then
        numberCount = numberCount + 1
        ReDim Preserve numberArray(UBound(numberArray) + 1)
        numberArray(UBound(numberArray)) = cell.Address
    End If
Next

Range("O18") = numberCount

If numberCount > 0 Then
    Dim Destination1 As Range
    Set Destination1 = Range("P18")
    Set Destination1 = Destination1.Resize(1, UBound(numberArray) + 1)
    Destination1.Value = numberArray
End If

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

