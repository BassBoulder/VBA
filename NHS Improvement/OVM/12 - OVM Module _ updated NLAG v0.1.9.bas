Attribute VB_Name = "Module21"
Sub aStart()

For I = 2 To 2
Dim Whitespace() As String
Dim Characters As String

''TESTS FOR WHITESPACE EXCLUSION FOR PHONE COLUMN *(FOR POST QUERY PHONE CHECKING)*

    ''COUNTS WHITESPACE IN CELL
    Whitespace = (Split(Range("X" & I), " "))
    MsgBox UBound(Whitespace)
    
    ''COUNTS CHARACTERS IN CELL
    Characters = (Len(Range("X" & I)))
    MsgBox Characters
    
    ''CONDENSED QUERY
    MsgBox UBound(Split(Range("X" & I), " "))

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
    'Call DQ_PIVOT
    
    
    'PIVOT Colour
    ActiveWorkbook.Sheets("PIVOT").Tab.ThemeColor = xlThemeColorAccent1
    
    'DQ PIVOT Colour
    'ActiveWorkbook.Sheets("DQ PIVOT").Tab.ThemeColor = xlThemeColorAccent2

    Sheets("PIVOT").Select
    Range("A1").Select
    
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

    ActiveWorkbook.Save

MsgBox "Completed - If you double click the 'Total' figure you will see that unique list of patients."


End Sub

Sub DeleteBlankRows()

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001

        If Range("A" & I).Value = "" Then Range("B" & I).ClearFormats
        
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
        , Origin:=xlWindows, StartRow:=2, DataType:=xlDelimited, TextQualifier _
        :=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:= _
        False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array _
        (6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), _
        Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array( _
        19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), _
        Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array( _
        32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), _
        Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array( _
        45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1), Array(50, 1), Array(51, 1), _
        Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), Array(56, 1), Array(57, 1), Array( _
        58, 1), Array(59, 1), Array(60, 1), Array(61, 1), Array(62, 1), Array(63, 1), Array(64, 1)), _
        TrailingMinusNumbers:=True
        
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
    
    ''DATE NOT RECOGNISED FIX
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D:D"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("D1").Value = "DATE_OF_BIRTH"

    Columns("H:H").Select
    Selection.TextToColumns Destination:=Range("H:H"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("H1").Value = "ALLOC_DATE"

    Columns("U:U").Select
    Selection.TextToColumns Destination:=Range("U:U"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("U1").Value = "ADD_BEF_DATE"
        
    Columns("N:N").Select
    Selection.TextToColumns Destination:=Range("N:N"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
    
        Range("N1").Value = "DATE_OF_DEATH"
    
    Columns("AA:AA").Select
    Selection.TextToColumns Destination:=Range("AA:AA"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("AA1").Value = "HO_END_DATE"
    
    Columns("AG:AG").Select
    Selection.TextToColumns Destination:=Range("AG:AG"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("AG1").Value = "OVM_SEEN_DATE"

    Application.DisplayAlerts = False

    Columns("AJ:AJ").Select
    Selection.TextToColumns Destination:=Range("AJ:AJ"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("AJ1").Value = ("OVM_DATE" & " " & "APPLIED") ''MISSING _ SO HAD TO DO THIS

        Range("AK1").Value = "EHIC_ID_NO"
        
     Application.DisplayAlerts = True
    
    Columns("AL:AL").Select
    Selection.TextToColumns Destination:=Range("AL:AL"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
        
        Range("AL1").Value = "EHIC_EXP_DATE"
    
    Columns("AY:AY").Select
    Selection.TextToColumns Destination:=Range("AY:AY"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
    
        Range("AY1").Value = "PRC_EHIC_EXP_DATE"
    
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

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001

If Range("C" & I) <> "" Then

    'FLAGS FOR RESPONSE CODES NOT 00 OR 0
    If Range("E" & I) <> "00" Or "0" Then
        Range("A" & I) = "Response Code " & Range("E" & I).Value
    End If

End If

Next

End Sub


Sub Weightings_Desc()
    
Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001

If Range("C" & I) <> "" Then

'LIKELY FREE
If Range("B" & I) < 1 Or Range("B" & I) = "" Then _
    Range("A" & I) = "4 - Likely Free"

'SOME EVIDENCE CHARGEABLE
If Range("B" & I) > 0 And Range("B" & I) < 20 Then _
    Range("A" & I) = "3 - Some Evidence Chargeable"
    
'LIKELY CHARGEABLE
If Range("B" & I) > 20 And Range("B" & I) < 998 Then _
    Range("A" & I) = "1 - Likely Chargeable"
    
'LIKELY RECOVERABLE
If Range("B" & I) > 998 Then _
    Range("A" & I) = "2 - Likely Recoverable"
        
        
End If
        
Next

End Sub
Sub WeightingsMinusNine()
''With new columns added ranges have shifted, new columns were added after this vba

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001
Dim NegLight As Integer
Dim Medium As Integer
Dim Light As Integer
Dim Heavy As Integer
NegLight = -9
Medium = 100
Light = 1
Heavy = 999
Today = Now

'Application.Calculation = xlCalculationManual

If Range("C" & I) <> "" Then

    ''Product Type = "EUS" + Immigration Status = ILR + Healthcae Status = "TRUE"
    If (Range("AD" & I).Value = "EUS") And _
        (Range("AE" & I).Value = "ILR") And _
        (Range("BE" & I).Value = True) Then
        
                    Range("A" & I).Value = NegLight
        
    End If
    

    ''HO-Status=Green(01), in-date / HO-Status=Green(03)
    If ((Range("E" & I).Value = "01" Or Range("F" & I).Value = "01" _
        Or Range("E" & I).Value = "1" Or Range("F" & I).Value = "1") And _
        (Range("AA" & I).Value = "" Or Range("AA" & I).Value > Range("H" & I).Value)) _
            Or ((Range("E" & I).Value = "03" Or Range("F" & I).Value = "03" _
                Or Range("E" & I).Value = "3" Or Range("F" & I).Value = "3") And _
                (Range("AA" & I).Value = "" Or Range("AA" & I).Value > Range("H" & I).Value) _
                 And (Range("E" & I).Value <> "02" Or Range("F" & I).Value <> "02" _
                  Or Range("E" & I).Value <> "2" Or Range("F" & I).Value <> "2")) Then
                
                    Range("A" & I).Value = NegLight
        
    End If
    
    ''OVM-Status=Cat-A, OVM-Status=Cat-B, OVM-Status=Cat-E
    If Range("G" & I).Value = "A" Or Range("G" & I).Value = "B" _
        And ((Range("E" & I).Value <> "02" Or Range("F" & I).Value <> "02") _
            Or (Range("E" & I).Value <> "2" Or Range("F" & I).Value <> "2")) Then
        
            Range("A" & I).Value = NegLight
        
    End If
            
    ''SUPERSEDED_BY
    If Range("I" & I).Value <> "" And (Range("E" & I).Value <> "02" _
        Or Range("F" & I).Value <> "02" _
        Or Range("E" & I).Value <> "2" _
        Or Range("F" & I).Value <> "2") Then
    
            Range("A" & I).Value = NegLight
        
    End If
    
    ''HO-Status=Red(02)
    If Range("E" & I).Value = "02" _
        Or Range("E" & I).Value = "2" _
        Or Range("F" & I).Value = "02" _
        Or Range("F" & I).Value = "2" Then
    
            Range("A" & I).Value = Medium
            
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
   
''HO-Status=Green(01), expired
If (Range("E" & I).Value = "01" Or Range("E" & I).Value = "1") And Range("AA" & I).Value <> "" And _
    Range("AA" & I).Value < Range("H" & I).Value Then _
        Range("A" & I).Value = Medium
    
''Date_of_Death & Response code 06
If Range("N" & I).Value <> "" And (Range("B" & I).Value = 6 Or Range("B" & I).Value = "06") Then _
        Range("A" & I).Value = Medium
    
''OVM-Status=Cat-D
If Range("G" & I).Value = "D" Then _
        Range("A" & I).Value = Medium
        
''OVM-Status=Cat-E
If Range("G" & I).Value = "E" Then _
        Range("A" & I).Value = Medium

''OVM-Status=Cat-F
If Range("G" & I).Value = "F" Then _
        Range("A" & I).Value = Medium

''EHIC=Yes
If (Range("AM" & I).Value <> "" And Range("AM" & I).Value <> "None" _
    And Range("AM" & I).Value <> "none") Then _
        Range("A" & I).Value = Heavy

''PRC=Yes
If (Range("AX" & I).Value <> "" And Range("AX" & I).Value <> "None" _
    And Range("AX" & I).Value <> "none") Then _
        Range("A" & I).Value = Heavy
    
''S1=Yes
If (Range("AQ" & I).Value <> "" And Range("AQ" & I).Value <> "None" _
    And Range("AQ" & I).Value <> "none") Then _
        Range("A" & I).Value = Heavy

''S2=Yes
If (Range("AU" & I).Value <> "" And Range("AU" & I).Value <> "None" _
    And Range("AU" & I).Value <> "none") Then _
        Range("A" & I).Value = Heavy
                
''OVM-Status=Cat-C
If Range("G" & I).Value = "C" Then _
        Range("A" & I).Value = Heavy
        
    
If Range("A" & I) = "" Then
                
    ''Address=Missing
    If Range("O" & I).Value = "" And Range("P" & I).Value = "" Then
            Range("A" & I).Value = Light
            
    End If
End If
    
If Range("A" & I) = "" Then
    
    
    ''Postcode=ZZ
    If Left(Range("T" & I).Value, 2) = "ZZ" Then
            Range("A" & I).Value = Light
    
    End If
End If
    
If Range("A" & I) = "" Then
    
    ''GP=No
    If Range("W" & I).Value = "" Then
            Range("A" & I).Value = Light
    
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End If

Next

End Sub

Sub Weightings()
 ''With new columns added ranges have shifted, new columns were added after this vba

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001
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

If Range("C" & I) <> "" And (Range("A" & I) = "" Or Range("A" & I) = "0") Then
    
''Old at NHS-No assignment
If Range("D" & I).Value <> "" And DateDiff("YYYY", Range("D" & I).Value, Now) > 15 And _
    Left(Range("C" & I).Value, 1) = 7 Then _
        Range("A" & I).Value = Light
                        
''OVM-Status=DecisionPending
If Range("G" & I).Value = "P" Then _
        Range("A" & I).Value = Light
    
'''Date_of_Death & Response code 06
'If Range("N" & I).Value <> "" And (Range("B" & I).Value = 6 Or Range("B" & I).Value = "06") Then _
'        Range("A" & I).Value = Medium

''NHS-No=Missing
If Range("C" & I).Value = "" Then _
        Range("A" & I).Value = Minor
        
''EUS Status = "FALSE"
If Range("AD" & I).Value = "EUS" And Range("AE" & I).Value = "ILR" And _
    Range("BE" & I).Value = False Then _
            Range("A" & I).Value = Minor
    
''EUS Status = "TRUE"
If Range("AD" & I).Value = "EUS" And Range("AE" & I).Value = "LTR" And _
    Range("BE" & I).Value = True Then _
            Range("A" & I).Value = Minor
        
''HO-Status=Green(01), expired
'If (Range("E" & I).Value = "01" Or Range("E" & I).Value = "1") And Range("AA" & I).Value <> "" And _
'    Range("AA" & I).Value < Range("H" & I).Value Then _
'        Range("A" & I).Value = Medium
        
'''OVM-Status=Cat-D
'If Range("G" & I).Value = "D" Then _
'        Range("A" & I).Value = Medium
'
'''OVM-Status=Cat-E
'If Range("G" & I).Value = "E" Then _
'        Range("A" & I).Value = Medium
'
'''OVM-Status=Cat-F
'If Range("G" & I).Value = "F" Then _
'        Range("A" & I).Value = Medium
'
'''EHIC=Yes
'If (Range("AM" & I).Value <> "" And Range("AM" & I).Value <> "None" _
'    And Range("AM" & I).Value <> "none") Then _
'        Range("A" & I).Value = Heavy
'
'''PRC=Yes
'If (Range("AX" & I).Value <> "" And Range("AX" & I).Value <> "None" _
'    And Range("AX" & I).Value <> "none") Then _
'        Range("A" & I).Value = Heavy
'
'''S1=Yes
'If (Range("AQ" & I).Value <> "" And Range("AQ" & I).Value <> "None" _
'    And Range("AQ" & I).Value <> "none") Then _
'        Range("A" & I).Value = Heavy
'
'''S2=Yes
'If (Range("AU" & I).Value <> "" And Range("AU" & I).Value <> "None" _
'    And Range("AU" & I).Value <> "none") Then _
'        Range("A" & I).Value = Heavy
'
'''OVM-Status=Cat-C
'If Range("G" & I).Value = "C" Then _
'        Range("A" & I).Value = Heavy
                
End If

Next

'Application.Calculation = xlCalculationAutomatic

End Sub

Sub WeightingsMinusNine_Rich_Desc()
''With new columns added ranges have shifted, new columns were added after this vba

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001
Dim NegLight As Integer
Dim Medium As Integer
NegLight = -9
Medium = 100


'Application.Calculation = xlCalculationManual

If Range("E" & I) <> "" Then

    ''Product Type = "EUS" + Immigration Status = ILR + Healthcare Status = "TRUE"
    If (Range("AF" & I).Value = "EUS" And Range("AG" & I).Value = "ILR" And _
        Range("BG" & I).Value = True) Then
        
            Range("A" & I).Value = "EUSS Status = TRUE"
    
    End If


    ''HO-Status=Green(01), in-date / HO-Status=Green(03)
    If ((Range("G" & I).Value = "01" Or Range("G" & I).Value = "1" Or Range("H" & I).Value = "01" _
        Or Range("H" & I).Value = "1") And _
            (Range("AB" & I).Value = "" Or Range("AB" & I).Value > Range("J" & I).Value)) _
            Or ((Range("G" & I).Value = "03" Or Range("G" & I).Value = "3" Or Range("H" & I).Value = "03" _
                Or Range("H" & I).Value = "3") And _
                (Range("AB" & I).Value = "" Or Range("AB" & I).Value > Range("J" & I).Value) _
                 And (Range("G" & I).Value <> "02" Or Range("G" & I).Value <> "2" Or Range("H" & I).Value <> "02" _
                    Or Range("H" & I).Value <> "2")) Then
                
                    Range("A" & I).Value = "HO-Status=Green(01), in-date / HO-Status=Green(03)"
        
    End If
    
    ''OVM-Status=Cat-A, OVM-Status=Cat-B, OVM-Status=Cat-E
    If Range("G" & I).Value = "A" Or Range("G" & I).Value = "B" _
        And (Range("E" & I).Value <> "02" Or Range("F" & I).Value <> "02" _
            Or Range("E" & I).Value <> "2" Or Range("F" & I).Value <> "2") Then
        
            Range("A" & I).Value = "OVM-Status=Cat-A, OVM-Status=Cat-B, OVM-Status=Cat-E"
        
    End If
            
    ''SUPERSEDED_BY
    If Range("K" & I).Value <> "" And (Range("G" & I).Value <> "02" _
        Or Range("H" & I).Value <> "02" Or Range("G" & I).Value <> "2" _
        Or Range("H" & I).Value <> "2") Then
    
            Range("A" & I).Value = "SUPERSEDED_BY"
        
    End If
    
    ''HO-Status=Red(02)
    If Range("G" & I).Value = "02" _
        Or Range("G" & I).Value = "2" _
        Or Range("H" & I).Value = "02" _
        Or Range("H" & I).Value = "2" Then
    
            Range("A" & I).Value = "HO-Status=Red(02)"
            
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
                
''Date_of_Death & Response code 06
If Range("J" & I).Value <> "" And (Range("D" & I).Value = 6 Or Range("D" & I).Value = "06") Then _
        Range("A" & I).Value = "Date_of_Death & Response code 06"
                
''HO-Status=Green(01), expired
If (Range("G" & I).Value = "01" Or Range("G" & I).Value = "1") _
    And Range("AC" & I).Value <> "" And _
        Range("AC" & I).Value < Range("J" & I).Value Then _
            Range("A" & I).Value = "HO-Status=Green(01), expired"
        
''OVM-Status=Cat-D
If Range("I" & I).Value = "D" Then _
        Range("A" & I).Value = "OVM-Status=Cat-D"
        
''OVM-Status=Cat-E
If Range("I" & I).Value = "E" Then _
        Range("A" & I).Value = "OVM-Status=Cat-E"

''OVM-Status=Cat-F
If Range("I" & I).Value = "F" Then _
        Range("A" & I).Value = "OVM-Status=Cat-F"

''EHIC=Yes
If (Range("AO" & I).Value <> "" And Range("AO" & I).Value <> "None" _
    And Range("AO" & I).Value <> "none") Then _
        Range("A" & I).Value = "EHIC=Yes"

''PRC=Yes
If (Range("AZ" & I).Value <> "" And Range("AZ" & I).Value <> "None" _
    And Range("AZ" & I).Value <> "none") Then _
        Range("A" & I).Value = "PRC=Yes"
    
''S1=Yes
If (Range("AS" & I).Value <> "" And Range("AS" & I).Value <> "None" _
    And Range("AS" & I).Value <> "none") Then _
        Range("A" & I).Value = "S1=Yes"

''S2=Yes
If (Range("AW" & I).Value <> "" And Range("AW" & I).Value <> "None" _
    And Range("AW" & I).Value <> "none") Then _
        Range("A" & I).Value = "S2=Yes"
        
''OVM-Status=Cat-C
If Range("I" & I).Value = "C" Then _
        Range("A" & I).Value = "OVM-Status=Cat-C"
                
                
                
If Range("A" & I) = "" Then
       
    ''Address=Missing
    If Range("Q" & I).Value = "" And Range("R" & I).Value = "" Then
            Range("A" & I).Value = "Address=Missing"
            
    End If
End If
    
    
If Range("A" & I) = "" Then
    
    ''Postcode=ZZ
    If Left(Range("V" & I).Value, 2) = "ZZ" Then
            Range("A" & I).Value = "Postcode=ZZ"
    
    End If
End If
    
    
If Range("A" & I) = "" Then
    
    ''GP=No
    If Range("Y" & I).Value = "" Then
            Range("A" & I).Value = "GP=No"
    
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End If

Next

End Sub

Sub Weightings_Rich_Desc()
 ''With new columns added ranges have shifted, new columns were added after this vba

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To 1001
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

If Range("E" & I) <> "" And (Range("A" & I) = "" Or Range("A" & I) = "0") Then

''EUS Status = "FALSE"
If Range("AF" & I).Value = "EUS" And Range("AG" & I).Value = "ILR" And _
    Range("BG" & I).Value = False Then _
            Range("A" & I).Value = "EUS Status = FALSE"
    
''EUS Status = "TRUE"
If Range("AF" & I).Value = "EUS" And Range("AG" & I).Value = "LTR" And _
    Range("BG" & I).Value = True Then _
            Range("A" & I).Value = "EUS Status = TRUE"

''Old at NHS-No assignment
If Range("F" & I).Value <> "" And DateDiff("YYYY", Range("F" & I).Value, Now) > 15 And _
    Left(Range("E" & I).Value, 1) = 7 Then _
        Range("A" & I).Value = "Old at NHS-No assignment"

''OVM-Status=DecisionPending
If Range("I" & I).Value = "P" Then _
        Range("A" & I).Value = "OVM-Status=DecisionPending"
    
''Date_of_Death & Response code 06
'If Range("J" & I).Value <> "" And (Range("D" & I).Value = 6 Or Range("D" & I).Value = "06") Then _
'        Range("A" & I).Value = "Date_of_Death & Response code 06"

''NHS-No=Missing
If Range("E" & I).Value = "" Then _
        Range("A" & I).Value = "NHS-No=Missing"
    
'''HO-Status=Green(01), expired
'If (Range("G" & I).Value = "01" Or Range("G" & I).Value = "1") _
'    And Range("AC" & I).Value <> "" And _
'        Range("AC" & I).Value < Range("J" & I).Value Then _
'            Range("A" & I).Value = "HO-Status=Green(01), expired"
'
'''OVM-Status=Cat-D
'If Range("I" & I).Value = "D" Then _
'        Range("A" & I).Value = "OVM-Status=Cat-D"
'
'''OVM-Status=Cat-E
'If Range("I" & I).Value = "E" Then _
'        Range("A" & I).Value = "OVM-Status=Cat-E"
'
'''OVM-Status=Cat-F
'If Range("I" & I).Value = "F" Then _
'        Range("A" & I).Value = "OVM-Status=Cat-F"
'
'''EHIC=Yes
'If (Range("AO" & I).Value <> "" And Range("AO" & I).Value <> "None" _
'    And Range("AO" & I).Value <> "none") Then _
'        Range("A" & I).Value = "EHIC=Yes"
'
'''PRC=Yes
'If (Range("AZ" & I).Value <> "" And Range("AZ" & I).Value <> "None" _
'    And Range("AZ" & I).Value <> "none") Then _
'        Range("A" & I).Value = "PRC=Yes"
'
'''S1=Yes
'If (Range("AS" & I).Value <> "" And Range("AS" & I).Value <> "None" _
'    And Range("AS" & I).Value <> "none") Then _
'        Range("A" & I).Value = "S1=Yes"
'
'''S2=Yes
'If (Range("AW" & I).Value <> "" And Range("AW" & I).Value <> "None" _
'    And Range("AW" & I).Value <> "none") Then _
'        Range("A" & I).Value = "S2=Yes"
'
'''OVM-Status=Cat-C
'If Range("I" & I).Value = "C" Then _
'        Range("A" & I).Value = "OVM-Status=Cat-C"
               
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

Dim FinalRow As Long
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

''COLOUR FORMATTING FOR RANKS
    
    ''-100 to 0
    Range("A2:A" & FinalRow).Select
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
    Range("A2:A" & FinalRow).Select
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
    Range("A2:A" & FinalRow).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=100", Formula2:="=200"
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
    
    ''201 to 10000
    Range("A2:A" & FinalRow).Select
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
    Range("A1:BN1001").Select
    
    Set objTable = Sheets(1).PivotTableWizard
    
    Set objfield = objTable.PivotFields("Weighting_Description")
    objfield.Orientation = xlRowField
    
    Set objfield = objTable.PivotFields("RESPONSE_CODE")
    objfield.Orientation = xlDataField
    objfield.Function = xlCount
    
    ''FIND THE BLANK ROW(S)
    ''With ActiveSheet.PivotTables(1).PivotFields("Weighting_Description" _
    ''    )
    ''    .PivotItems("(blank)").Visible = False
    ''End With
    
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

    

End Sub

Sub DQ_PIVOT()

''PIVOT CREATOR

Dim pt As PivotTable
Dim rf As PivotField
Dim Total As Integer

    'RENAMING DATA SHEET
    Sheets("DATA").Select
    Range("A1:BN1001").Select
    
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
