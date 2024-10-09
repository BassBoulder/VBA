Attribute VB_Name = "Module21"
Sub aStart()

For i = 2 To 2
Dim Whitespace() As String
Dim Characters As String

''TESTS FOR WHITESPACE EXCLUSION FOR PHONE COLUMN *(FOR POST QUERY PHONE CHECKING)*

    ''COUNTS WHITESPACE IN CELL
    Whitespace = (Split(Range("X" & i), " "))
    MsgBox UBound(Whitespace)
    
    ''COUNTS CHARACTERS IN CELL
    Characters = (Len(Range("X" & i)))
    MsgBox Characters
    
    ''CONDENSED QUERY
    MsgBox UBound(Split(Range("X" & i), " "))

Next

End Sub

Sub BBBRunInOrder()
Attribute BBBRunInOrder.VB_ProcData.VB_Invoke_Func = "B\n14"

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

''RUNS ALL CODES IN SEQUENCE
    Call Open_Dat_advanced
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
    Call Insert_Weighting
    Call WeightingsMinusNine
    Call Weightings
    Call Colour_Formatting
    Call Insert_Weighting_Desc
    Call Weightings_Desc

    Call Insert_Weighting_Rich_Desc
    Call WeightingsMinusNine_Rich_Desc
    Call Weightings_Rich_Desc
    Call Insert_DQ_Issues
    Call DQ_Desc

    Call Filtering

        Cells.Select
        Cells.EntireColumn.AutoFit
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

    Call DeleteBlankRows

    Call PIVOT
    Call PIVOT_Formatting
    Call DQ_PIVOT
    
    
    'PIVOT Colour
    ActiveWorkbook.Sheets("PIVOT").Tab.ThemeColor = xlThemeColorAccent1
    
    'DQ PIVOT Colour
    ActiveWorkbook.Sheets("DQ PIVOT").Tab.ThemeColor = xlThemeColorAccent2

    Sheets("PIVOT").Select
    Range("A1").Select
    
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

    ActiveWorkbook.Save

MsgBox "Completed - If you double click the 'Total' figure you will see that unique list of patients."


End Sub

Sub DeleteBlankRows()

For i = 2 To 5000

        If Range("A" & i).Value = "" Then Range("B" & i).ClearFormats
        
Next

End Sub


Sub Open_Dat_advanced()

 Dim fd As Office.FileDialog

 Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd

   .AllowMultiSelect = False
   .InitialFileName = ThisWorkbook.Path & "\" _

   ' Set the title of the dialog box.
   .Title = "Please select the file."

   ' Clear out the current filters, and add our own.
   .Filters.Clear
   '.Filters.Add "Excel 2003", "*.xls"
   .Filters.Add "All Files", "*.*"

   ' Show the dialog box. If the .Show method returns True, the
   ' user picked at least one file. If the .Show method returns
   ' False, the user clicked Cancel.
   If .Show = True Then
     txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox


''OPENs .DAT FILES
    Workbooks.OpenText Filename:= _
        txtFileName _
        , Origin:=xlMSDOS, StartRow:=2, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=True, OtherChar:="|", FieldInfo:= _
        Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 2), Array(5, 2), Array(6, 2), Array(7 _
        , 4), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array _
        (14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 4), _
        Array(21, 2), Array(22, 1), Array(23, 4), Array(24, 2), Array(25, 2), Array(26, 4), Array( _
        27, 1), Array(28, 1), Array(29, 2), Array(30, 2), Array(31, 2), Array(32, 4), Array(33, 2), _
        Array(34, 2), Array(35, 1), Array(36, 2), Array(37, 2), Array(38, 4), Array(39, 2), Array( _
        40, 2), Array(41, 2), Array(42, 2), Array(43, 2), Array(44, 2), Array(45, 2), Array(46, 2), _
        Array(47, 2), Array(48, 2), Array(49, 2), Array(50, 2), Array(51, 1), Array(52, 2), Array( _
        53, 2), Array(54, 2), Array(55, 2)), TrailingMinusNumbers:=True

   End If
   
    SaveChoice = InputBox("File Name?", "Save As", "")

    ActiveWorkbook.SaveAs Filename:= _
    ThisWorkbook.Path & "\" & SaveChoice _
    , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        

End With

End Sub
Sub Insert_Weighting()

''INSERTS A WEIGHTING SCORING COLUMN
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Weighting"

End Sub
Sub Insert_Weighting_Desc()

''INSERTS A WEIGHTING SCORING DESCRIPTION
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Weighting_Description"
    Columns("A:BF").EntireColumn.AutoFit

End Sub

Sub Insert_Weighting_Rich_Desc()

''INSERTS A WEIGHTING SCORING RICH DESCRIPTION COLUMN
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Weighting_Rich_Description"

End Sub

Sub Insert_DQ_Issues()

''INSERTS A DATA QUALITY SCORING COLUMN
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Data Quality Issues"

End Sub

Sub DQ_Desc()

For i = 2 To 5000

If Range("C" & i) <> "" Then

    'FLAGS FOR RESPONSE CODES NOT 00 OR 0
    If Range("E" & i) <> "00" Or "0" Then
        Range("A" & i) = "Response Code " & Range("E" & i).Value
    End If

End If

Next

End Sub


Sub Weightings_Desc()
    
For i = 2 To 5000

If Range("C" & i) <> "" Then

'LIKELY FREE
If Range("B" & i) < 1 Or Range("B" & i) = "" Then _
    Range("A" & i) = "Likely Free"

'SOME EVIDENCE CHARGEABLE
If Range("B" & i) > 0 And Range("B" & i) < 20 Then _
    Range("A" & i) = "Some Evidence Chargeable"
    
'LIKELY CHARGEABLE
If Range("B" & i) > 20 And Range("B" & i) < 998 Then _
    Range("A" & i) = "Likely Chargeable"
    
'LIKELY RECOVERABLE
If Range("B" & i) > 998 Then _
    Range("A" & i) = "Likely Recoverable"
        
        
End If
        
Next

End Sub
Sub WeightingsMinusNine()
''With new columns added ranges have shifted, new columns were added after this vba

For i = 2 To 5000
Dim NegLight As Integer
Dim Medium As Integer
Dim Light As Integer
NegLight = -9
Medium = 100
Light = 1


'Application.Calculation = xlCalculationManual

If Range("C" & i) <> "" Then

    ''HO-Status=Green(01), in-date / HO-Status=Green(03)
    If ((Range("E" & i).Value = "01" Or Range("F" & i).Value = "01") And _
        (Range("AA" & i).Value = "" Or Range("AA" & i).Value > Range("H" & i).Value)) _
            Or ((Range("E" & i).Value = "03" Or Range("F" & i).Value = "03") And _
                (Range("AA" & i).Value = "" Or Range("AA" & i).Value > Range("H" & i).Value) _
                 And (Range("E" & i).Value <> "02" Or Range("F" & i).Value <> "02")) Then
                
                    Range("A" & i).Value = NegLight
        
    End If
    
    ''OVM-Status=Cat-A, OVM-Status=Cat-B, OVM-Status=Cat-E
    If Range("G" & i).Value = "A" Or Range("G" & i).Value = "B" _
        And (Range("E" & i).Value <> "02" Or Range("F" & i).Value <> "02") Then
        
            Range("A" & i).Value = NegLight
        
    End If
            
    ''SUPERSEDED_BY
    If Range("I" & i).Value <> "" And (Range("E" & i).Value <> "02" _
        Or Range("F" & i).Value <> "02") Then
    
            Range("A" & i).Value = NegLight
        
    End If
    
    ''HO-Status=Red(02)
    If Range("E" & i).Value = "02" _
        Or Range("E" & i).Value = "02" _
        Or Range("F" & i).Value = "02" _
        Or Range("F" & i).Value = "02" Then
    
            Range("A" & i).Value = Medium
            
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''If Range("A" & i) = "" Then
    
    
    ''''Overseas Tel No=Yes ''(Covers the fact some numbers have encapsulation i.e. (01482...
    ''If Len(Range("V" & i)) <> ("6" And "7") And _
    ''    (Len(Range("V" & i)) <> ("8" And (UBound(Split(Range("V" & i), " ")) = 1))) _
    ''        And Range("V" & i).Value <> "" Then
    ''
    ''    If (Left(Range("V" & i).Value, 2) <> "01" And Left(Range("V" & i).Value, 3) <> "(01") And _
    ''        (Left(Range("V" & i).Value, 2) <> "02" And Left(Range("V" & i).Value, 3) <> "(02") And _
    ''        (Left(Range("V" & i).Value, 2) <> "03" And Left(Range("V" & i).Value, 3) <> "(03") And _
    ''        (Left(Range("V" & i).Value, 4) <> "0800" And Left(Range("V" & i).Value, 5) <> "(0800") And _
    ''        (Left(Range("V" & i).Value, 2) <> "07" And Left(Range("V" & i).Value, 3) <> "(07") And _
    ''        (Left(Range("V" & i).Value, 4) <> "0808" And Left(Range("V" & i).Value, 5) <> "(0808") And _
    ''        (Left(Range("V" & i).Value, 3) <> "084" And Left(Range("V" & i).Value, 4) <> "(084") And _
    ''        (Left(Range("V" & i).Value, 3) <> "087" And Left(Range("V" & i).Value, 4) <> "(087") And _
    ''        (Left(Range("V" & i).Value, 3) <> "09" And Left(Range("V" & i).Value, 3) <> "(09") And _
    ''        (Left(Range("V" & i).Value, 3) <> "+44" And Left(Range("V" & i).Value, 4) <> "(+44") Then _
    ''            Range("A" & i).Value = Light
    ''
    ''End If
   ''End If
    
If Range("A" & i) = "" Then
                
    ''Address=Missing
    If Range("O" & i).Value = "" And Range("P" & i).Value = "" Then
            Range("A" & i).Value = Light
            
    End If
End If
    
If Range("A" & i) = "" Then
    
    
    ''Postcode=ZZ
    If Left(Range("T" & i).Value, 2) = "ZZ" Then
            Range("A" & i).Value = Light
    
    End If
End If
    
If Range("A" & i) = "" Then
    
    ''GP=No
    If Range("W" & i).Value = "" Then
            Range("A" & i).Value = Light
    
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End If

Next

End Sub

Sub Weightings()
 ''With new columns added ranges have shifted, new columns were added after this vba

For i = 2 To 5000
Dim Light As Integer
Dim Minor As Integer
Dim Medium As Integer
Dim Heavy As Integer
Dim Today
Dim Whitespace() As String
Light = 1
Minor = 3
Medium = 100
Heavy = 999
Today = Now

If Range("C" & i) <> "" And (Range("A" & i) = "" Or Range("A" & i) = "0") Then
    
''Old at NHS-No assignment
If Range("D" & i).Value <> "" And DateDiff("YYYY", Range("D" & i).Value, Now) > 15 And _
    Left(Range("C" & i).Value, 1) = 7 Then _
        Range("A" & i).Value = Light
                        
''OVM-Status=DecisionPending
If Range("G" & i).Value = "P" Then _
        Range("A" & i).Value = Light
    
''Date_of_Death & Response code 06
If Range("N" & i).Value <> "" And Range("B" & i).Value = 6 Then _
        Range("A" & i).Value = Medium

''NHS-No=Missing
If Range("C" & i).Value = "" Then _
        Range("A" & i).Value = Minor
    
''HO-Status=Green(01), expired
If Range("E" & i).Value = "01" And Range("AA" & i).Value <> "" And _
    Range("AA" & i).Value < Range("H" & i).Value Then _
        Range("A" & i).Value = Medium
        
''OVM-Status=Cat-D
If Range("G" & i).Value = "D" Then _
        Range("A" & i).Value = Medium
        
''OVM-Status=Cat-E
If Range("G" & i).Value = "E" Then _
        Range("A" & i).Value = Medium

''OVM-Status=Cat-F
If Range("G" & i).Value = "F" Then _
        Range("A" & i).Value = Medium

''EHIC=Yes
If (Range("AM" & i).Value <> "" And Range("AM" & i).Value <> "None" _
    And Range("AM" & i).Value <> "none") Then _
        Range("A" & i).Value = Heavy

''PRC=Yes
If (Range("AX" & i).Value <> "" And Range("AX" & i).Value <> "None" _
    And Range("AX" & i).Value <> "none") Then _
        Range("A" & i).Value = Heavy
    
''S1=Yes
If (Range("AQ" & i).Value <> "" And Range("AQ" & i).Value <> "None" _
    And Range("AQ" & i).Value <> "none") Then _
        Range("A" & i).Value = Heavy

''S2=Yes
If (Range("AU" & i).Value <> "" And Range("AU" & i).Value <> "None" _
    And Range("AU" & i).Value <> "none") Then _
        Range("A" & i).Value = Heavy
                
''OVM-Status=Cat-C
If Range("G" & i).Value = "C" Then _
        Range("A" & i).Value = Heavy
                
End If

Next

'Application.Calculation = xlCalculationAutomatic

End Sub

Sub WeightingsMinusNine_Rich_Desc()
''With new columns added ranges have shifted, new columns were added after this vba

For i = 2 To 5000
Dim NegLight As Integer
Dim Medium As Integer
NegLight = -9
Medium = 100


'Application.Calculation = xlCalculationManual

If Range("E" & i) <> "" Then

    ''HO-Status=Green(01), in-date / HO-Status=Green(03)
    If ((Range("G" & i).Value = "01" Or Range("H" & i).Value = "01") And _
        (Range("AC" & i).Value = "" Or Range("AC" & i).Value > Range("J" & i).Value)) _
            Or ((Range("G" & i).Value = "03" Or Range("H" & i).Value = "03") And _
                (Range("AC" & i).Value = "" Or Range("AC" & i).Value > Range("J" & i).Value) _
                 And (Range("G" & i).Value <> "02" Or Range("H" & i).Value <> "02")) Then
                
                    Range("A" & i).Value = "HO-Status=Green(01), in-date / HO-Status=Green(03)"
        
    End If
    
    ''OVM-Status=Cat-A, OVM-Status=Cat-B, OVM-Status=Cat-E
    If Range("G" & i).Value = "A" Or Range("G" & i).Value = "B" _
        And (Range("E" & i).Value <> "02" Or Range("F" & i).Value <> "02") Then
        
            Range("A" & i).Value = "OVM-Status=Cat-A, OVM-Status=Cat-B, OVM-Status=Cat-E"
        
    End If
            
    ''SUPERSEDED_BY
    If Range("K" & i).Value <> "" And (Range("G" & i).Value <> "02" _
        Or Range("H" & i).Value <> "02") Then
    
            Range("A" & i).Value = "SUPERSEDED_BY"
        
    End If
    
    ''HO-Status=Red(02)
    If Range("G" & i).Value = "02" _
        Or Range("G" & i).Value = "02" _
        Or Range("H" & i).Value = "02" _
        Or Range("H" & i).Value = "02" Then
    
            Range("A" & i).Value = "HO-Status=Red(02)"
            
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If Range("A" & i) = "" Then
    
    
    ''''Overseas Tel No=Yes ''(Covers the fact some numbers have encapsulation i.e. (01482...
    ''If Len(Range("X" & i)) <> ("6" And "7") And _
    ''    (Len(Range("X" & i)) <> ("8" And (UBound(Split(Range("X" & i), " ")) = 1))) _
    ''        And Range("X" & i).Value <> "" Then
    ''
    ''If (Left(Range("X" & i).Value, 2) <> "01" And Left(Range("X" & i).Value, 3) <> "(01") And _
    ''    (Left(Range("X" & i).Value, 2) <> "02" And Left(Range("X" & i).Value, 3) <> "(02") And _
    ''    (Left(Range("X" & i).Value, 2) <> "03" And Left(Range("X" & i).Value, 3) <> "(03") And _
    ''    (Left(Range("X" & i).Value, 4) <> "0800" And Left(Range("X" & i).Value, 5) <> "(0800") And _
    ''    (Left(Range("X" & i).Value, 2) <> "07" And Left(Range("X" & i).Value, 3) <> "(07") And _
    ''    (Left(Range("X" & i).Value, 4) <> "0808" And Left(Range("X" & i).Value, 5) <> "(0808") And _
    ''    (Left(Range("X" & i).Value, 3) <> "084" And Left(Range("X" & i).Value, 4) <> "(084") And _
    ''    (Left(Range("X" & i).Value, 3) <> "087" And Left(Range("X" & i).Value, 4) <> "(087") And _
    ''    (Left(Range("X" & i).Value, 3) <> "09" And Left(Range("X" & i).Value, 3) <> "(09") And _
    ''    (Left(Range("X" & i).Value, 3) <> "+44" And Left(Range("X" & i).Value, 4) <> "(+44") Then _
    ''        Range("A" & i).Value = "Overseas Tel No=Yes"
    ''
    ''End If
  ''End If
                
                
If Range("A" & i) = "" Then
       
    ''Address=Missing
    If Range("Q" & i).Value = "" And Range("R" & i).Value = "" Then
            Range("A" & i).Value = "Address=Missing"
            
    End If
End If
    
    
If Range("A" & i) = "" Then
    
    ''Postcode=ZZ
    If Left(Range("V" & i).Value, 2) = "ZZ" Then
            Range("A" & i).Value = "Postcode=ZZ"
    
    End If
End If
    
    
If Range("A" & i) = "" Then
    
    ''GP=No
    If Range("Y" & i).Value = "" Then
            Range("A" & i).Value = "GP=No"
    
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End If

Next

End Sub

Sub Weightings_Rich_Desc()
 ''With new columns added ranges have shifted, new columns were added after this vba

For i = 2 To 5000
Dim Light As Integer
Dim Minor As Integer
Dim Medium As Integer
Dim Heavy As Integer
Dim Today
Light = 1
Minor = 3
Medium = 100
Heavy = 999
Today = Now

If Range("E" & i) <> "" And (Range("A" & i) = "" Or Range("A" & i) = "0") Then

''Old at NHS-No assignment
If Range("F" & i).Value <> "" And DateDiff("YYYY", Range("F" & i).Value, Now) > 15 And _
    Left(Range("E" & i).Value, 1) = 7 Then _
        Range("A" & i).Value = "Old at NHS-No assignment"

''OVM-Status=DecisionPending
If Range("I" & i).Value = "P" Then _
        Range("A" & i).Value = "OVM-Status=DecisionPending"
    
''Date_of_Death & Response code 06
If Range("J" & i).Value <> "" And Range("D" & i).Value = 6 Then _
        Range("A" & i).Value = "Date_of_Death & Response code 06"

''NHS-No=Missing
If Range("E" & i).Value = "" Then _
        Range("A" & i).Value = "NHS-No=Missing"
    
''HO-Status=Green(01), expired
If Range("G" & i).Value = "01" And Range("AC" & i).Value <> "" And _
    Range("AC" & i).Value < Range("J" & i).Value Then _
        Range("A" & i).Value = "HO-Status=Green(01), expired"
        
''OVM-Status=Cat-D
If Range("I" & i).Value = "D" Then _
        Range("A" & i).Value = "OVM-Status=Cat-D"
        
''OVM-Status=Cat-E
If Range("I" & i).Value = "E" Then _
        Range("A" & i).Value = "OVM-Status=Cat-E"

''OVM-Status=Cat-F
If Range("I" & i).Value = "F" Then _
        Range("A" & i).Value = "OVM-Status=Cat-F"

''EHIC=Yes
If (Range("AO" & i).Value <> "" And Range("AO" & i).Value <> "None" _
    And Range("AO" & i).Value <> "none") Then _
        Range("A" & i).Value = "EHIC=Yes"

''PRC=Yes
If (Range("AZ" & i).Value <> "" And Range("AZ" & i).Value <> "None" _
    And Range("AZ" & i).Value <> "none") Then _
        Range("A" & i).Value = "PRC=Yes"
    
''S1=Yes
If (Range("AS" & i).Value <> "" And Range("AS" & i).Value <> "None" _
    And Range("AS" & i).Value <> "none") Then _
        Range("A" & i).Value = "S1=Yes"

''S2=Yes
If (Range("AW" & i).Value <> "" And Range("AW" & i).Value <> "None" _
    And Range("AW" & i).Value <> "none") Then _
        Range("A" & i).Value = "S2=Yes"
        
''OVM-Status=Cat-C
If Range("I" & i).Value = "C" Then _
        Range("A" & i).Value = "OVM-Status=Cat-C"
               
End If

Next

'Application.Calculation = xlCalculationAutomatic

End Sub


Sub Filtering()

''ADDS FITLER AND ORDERS BY HIGHEST FIRST
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Range("A2").Select
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub
Sub Colour_Formatting()

''COLOUR FORMATTING FOR RANKS
    
    ''-100 to 0
    Range("A2:A5000").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=-100", Formula2:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    ''1 to 99
    Range("A2:A5000").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=1", Formula2:="=99"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ''100 to 200
    Range("A2:A5000").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=100", Formula2:="=200"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ''201 to 10000
    Range("A2:A5000").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=201", Formula2:="=10000"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Sub PIVOT()

''PIVOT CREATOR

Dim pt As PivotTable
Dim rf As PivotField
Dim Total As Integer

    'RENAMING DATA SHEET
    ActiveSheet.Name = "DATA"
    Range("A1:BN50001").Select
    
    Set objTable = Sheets(1).PivotTableWizard
    
    Set objfield = objTable.PivotFields("Weighting_Description")
    objfield.Orientation = xlRowField
    
    Set objfield = objTable.PivotFields("RESPONSE_CODE")
    objfield.Orientation = xlDataField
    objfield.Function = xlCount
    
    ''FIND THE BLANK ROW(S)
    With ActiveSheet.PivotTables(1).PivotFields("Weighting_Description" _
        )
        .PivotItems("(blank)").Visible = False
    End With
    
    'RENAMING PIVOT SHEET
    ActiveSheet.Name = "PIVOT"
    
    With ActiveSheet.PivotTables(1).PivotFields("Weighting_Rich_Description")
        .Orientation = xlRowField
        .Position = 2
    End With
    Range("A6").Select
    ActiveSheet.PivotTables(1).PivotFields("Weighting_Description"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        
    'TIDYING UP FORMAT
    Columns("A:C").EntireColumn.AutoFit
    
    'HIDE FIELD LIST
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    'COLOUR ORDER
    Total = (ActiveSheet.PivotTables(1).PivotFields("Weighting_Description").PivotItems.Count - 1)
    
    'Testing tool to pick order
    'MsgBox (Total)
    'Total = InputBox("pick a number", "DO IT!")


    'SAFETY MEASURE "IF" TO STOP A BREAK IF NO "Likely Free" CATEGORY IS AVAILABLE
    ActiveSheet.PivotTables(1).PivotFields("Weighting_Description"). _
    PivotItems("Likely Free").Position = Total

End Sub

Sub DQ_PIVOT()

''PIVOT CREATOR

Dim pt As PivotTable
Dim rf As PivotField
Dim Total As Integer

    'RENAMING DATA SHEET
    Sheets("DATA").Select
    Range("A1:BN50001").Select
    
    Set objTable = Sheets(2).PivotTableWizard
    
    Set objfield = objTable.PivotFields("Data Quality Issues")
    objfield.Orientation = xlRowField
    
    Set objfield = objTable.PivotFields("Data Quality Issues")
    objfield.Orientation = xlDataField
    objfield.Function = xlCount
    
    ''FIND THE BLANK ROW(S)
    
    If Range("A3") <> "(blank)" Then
    
         With ActiveSheet.PivotTables(1).PivotFields("Data Quality Issues" _
            )
            .PivotItems("(blank)").Visible = False
        End With
    
    End If
    
    'RENAMING PIVOT SHEET
    ActiveSheet.Name = "DQ PIVOT"
    
    With ActiveSheet.PivotTables(1).PivotFields("Data Quality Issues")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("A6").Select
    ActiveSheet.PivotTables(1).PivotFields("Data Quality Issues"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        
    'TIDYING UP FORMAT
    Columns("A:C").EntireColumn.AutoFit
    
    'HIDE FIELD LIST
    ActiveWorkbook.ShowPivotTableFieldList = False
    
End Sub
Sub PIVOT_Formatting()

'FORMATS THE PIVOT WITH COLOUR

    'RANGE FOR CONDITIONAL FORMATTING
    Range("A1:A50").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Likely Free", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936 'GREEN
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Some Evidence chargeable", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407 'AMBER
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Likely Chargeable", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255 'RED
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Likely Recoverable", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255 'RED
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
