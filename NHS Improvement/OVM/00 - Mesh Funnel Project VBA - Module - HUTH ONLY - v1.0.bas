Attribute VB_Name = "Module1"
Public SizeCheck As Long

Sub AAARunInOrder()
Attribute AAARunInOrder.VB_ProcData.VB_Invoke_Func = "A\n14"

FailCheck = MsgBox("Has BI Failed?", vbYesNo)
SizeCheck = InputBox("What size do you want the batch in (default = 5000)", "Batch Size", 5000)

'Application.Calculation = xlCalculationManual
'Application.ScreenUpdating = False

''RUNS ALL CODES IN SEQUENCE
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Mesh_RAW"
    
    
''CHECK IF BI HAS FAILED
    If FailCheck = vbYes Then
            Call RunSQL_BIFail
        Else
            Call RunSQL
    End If

    Call Reformat_Date
    Call OpenOVM
    
    Call Copy_AND_Paste_alpha
    Call Sheet1.SaveExtractFile
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_beta
    Call Sheet1.SaveExtractFile
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_charlie
    Call Sheet1.SaveExtractFile
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_delta
    Call Sheet1.SaveExtractFile
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_echo
    Call Sheet1.SaveExtractFile
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_foxtrot
    Call Sheet1.SaveExtractFile 'Using existing vba
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))

    Call Copy_AND_Paste_gamma
    Call Sheet1.SaveExtractFile 'Using existing vba
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_hotel
    Call Sheet1.SaveExtractFile 'Using existing vba
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_indigo
    Call Sheet1.SaveExtractFile 'Using existing vba
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Copy_AND_Paste_juliet
    Call Sheet1.SaveExtractFile 'Using existing vba
    
        'Artificial Slowdown due to Excel Export Issues
        Application.Wait (Now + TimeValue("0:00:01"))
    
    Call Cleanse
    
    Sheets("OVM Request").Select

'Application.Calculation = xlCalculationAutomatic
'Application.ScreenUpdating = True
    
    MsgBox "They'll be 1 file for every " & SizeCheck & " patients, going to a maximum of " & ((SizeCheck * 10) - 1) & "."

End Sub
Sub Cleanse()

Application.DisplayAlerts = False

'Cleaning Up, Data safety etc.
Sheets("Mesh_RAW").Delete
Sheets("NHS Numbers").Select
    Selection.ClearContents
    
Application.DisplayAlerts = True

End Sub

Sub RunSQL()

Dim User As String
Dim Host As String
User = (Environ$("Username"))
Host = Environ$("computername")
'UserName = InputBox("Is This You?", "Is This You?", User)

''CREATE ODBC CONNECTION + INJECT SQL + RUN SQL
Workbooks(1).Connections.AddFromFile _
    "C:\Users\" & User & "\Documents\My Data Sources\CHH-BILive HealthBI.odc"
    With ActiveWorkbook.Connections("HealthBI").OLEDBConnection
        .BackgroundQuery = True
        .CommandText = Array("DECLARE @START AS Date DECLARE @END AS Date SET @START =  CASE WHEN DATEPART(WEEKDAY,GETDATE()) = '1' THEN GETDATE" _
        , _
        "()-3 WHEN DATEPART(WEEKDAY,GETDATE()) IN ('2','3','4','5') THEN GETDATE()-2  END SET @END = GETDATE()-1 SELECT" _
        , _
        " mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_OP_REFERRAL ref INNER JOIN CDO_MPI mpi ON ref.CDO_MPI_UNIQUE_ID = mpi." _
        , _
        "UNIQUE_ID WHERE ref.EFFECTIVE_WAITING_START_DATE BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SE" _
        , _
        "LECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_APC_HOSPITAL_PROVIDER_SPELL adm INNER JOIN CDO_MPI mpi ON adm.CDO_" _
        , _
        "MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE adm.ADMISSION_METHOD_HOSPITAL_PROVIDER_SPELL NOT IN ('11','12','13') AND adm.ST" _
        , _
        "ART_DATE_HOSPITAL_PROVIDER_SPELL BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUM" _
        , _
        "BER, PERSON_BIRTH_DATE FROM CDO_OP_APPOINTMENT opa INNER JOIN CDO_MPI mpi ON opa.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID" _
        , _
        " WHERE (opa.APPOINTMENT_START_DATE BETWEEN @START AND @END OR opa.APPOINTMENT_BOOKED_DATE BETWEEN @START AND @END)" _
        , _
        " AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_A_AND_E_ATTENDANCE e" _
        , _
        "d INNER JOIN CDO_MPI mpi ON ed.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE ed.ARRIVAL_DATE BETWEEN @START AND @END AND " _
        , "NHS_NUMBER IS NOT NULL GROUP BY mpi.NHS_NUMBER, PERSON_BIRTH_DATE")
        .CommandType = xlCmdSql
        .Connection = Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=HealthBI;Data Source=CHH-BILive;Use Proced" _
        , _
        "ure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=Host;Use Encryption for Data=False;Tag with column col" _
        , "lation when possible=False")
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("HealthBI")
        .Name = "HealthBI"
        .Description = ""
    End With
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=HealthBI;Data Source=CHH-BILive;Use Proced" _
        , _
        "ure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=Host;Use Encryption for Data=False;Tag with column col" _
        , "lation when possible=False"), Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("DECLARE @START AS Date DECLARE @END AS Date SET @START =  CASE WHEN DATEPART(WEEKDAY,GETDATE()) = '1' THEN GETDATE" _
        , _
        "()-3 WHEN DATEPART(WEEKDAY,GETDATE()) IN ('2','3','4','5') THEN GETDATE()-2  END SET @END = GETDATE()-1 SELECT" _
        , _
        " mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_OP_REFERRAL ref INNER JOIN CDO_MPI mpi ON ref.CDO_MPI_UNIQUE_ID = mpi." _
        , _
        "UNIQUE_ID WHERE ref.EFFECTIVE_WAITING_START_DATE BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SE" _
        , _
        "LECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_APC_HOSPITAL_PROVIDER_SPELL adm INNER JOIN CDO_MPI mpi ON adm.CDO_" _
        , _
        "MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE adm.ADMISSION_METHOD_HOSPITAL_PROVIDER_SPELL NOT IN ('11','12','13') AND adm.ST" _
        , _
        "ART_DATE_HOSPITAL_PROVIDER_SPELL BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUM" _
        , _
        "BER, PERSON_BIRTH_DATE FROM CDO_OP_APPOINTMENT opa INNER JOIN CDO_MPI mpi ON opa.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID" _
        , _
        " WHERE (opa.APPOINTMENT_START_DATE BETWEEN @START AND @END OR opa.APPOINTMENT_BOOKED_DATE BETWEEN @START AND @END)" _
        , _
        " AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_A_AND_E_ATTENDANCE e" _
        , _
        "d INNER JOIN CDO_MPI mpi ON ed.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE ed.ARRIVAL_DATE BETWEEN @START AND @END AND " _
        , "NHS_NUMBER IS NOT NULL GROUP BY mpi.NHS_NUMBER, PERSON_BIRTH_DATE")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "CHH_BILive_HealthBI"
        .Refresh BackgroundQuery:=False
    End With
    

End Sub

Sub RunSQL_BIFail()

Dim User As String
Dim Host As String
User = (Environ$("Username"))
Host = Environ$("computername")
StartDate = InputBox("What is the Start Date?", "BI Failure Workaround - Start Date", "dd/mm/yyyy")
EndDate = InputBox("What is the End Date?", "BI Failure Workaround - End Date", "dd/mm/yyyy")

''CREATE ODBC CONNECTION + INJECT SQL + RUN SQL
Workbooks(1).Connections.AddFromFile _
    "C:\Users\" & User & "\Documents\My Data Sources\CHH-BILive HealthBI.odc"
    With ActiveWorkbook.Connections("HealthBI").OLEDBConnection
        .BackgroundQuery = True
        .CommandText = Array("DECLARE @START AS Date DECLARE @END AS Date SET @START =" & "'" & StartDate & "'" & " SET @END =" & "'" & EndDate & "'" & " SELECT" _
        , _
        " mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_OP_REFERRAL ref INNER JOIN CDO_MPI mpi ON ref.CDO_MPI_UNIQUE_ID = mpi." _
        , _
        "UNIQUE_ID WHERE ref.EFFECTIVE_WAITING_START_DATE BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SE" _
        , _
        "LECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_APC_HOSPITAL_PROVIDER_SPELL adm INNER JOIN CDO_MPI mpi ON adm.CDO_" _
        , _
        "MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE adm.ADMISSION_METHOD_HOSPITAL_PROVIDER_SPELL NOT IN ('11','12','13') AND adm.ST" _
        , _
        "ART_DATE_HOSPITAL_PROVIDER_SPELL BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUM" _
        , _
        "BER, PERSON_BIRTH_DATE FROM CDO_OP_APPOINTMENT opa INNER JOIN CDO_MPI mpi ON opa.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID" _
        , _
        " WHERE (opa.APPOINTMENT_START_DATE BETWEEN @START AND @END OR opa.APPOINTMENT_BOOKED_DATE BETWEEN @START AND @END)" _
        , _
        " AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_A_AND_E_ATTENDANCE e" _
        , _
        "d INNER JOIN CDO_MPI mpi ON ed.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE ed.ARRIVAL_DATE BETWEEN @START AND @END AND " _
        , "NHS_NUMBER IS NOT NULL GROUP BY mpi.NHS_NUMBER, PERSON_BIRTH_DATE")
        .CommandType = xlCmdSql
        .Connection = Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=HealthBI;Data Source=CHH-BILive;Use Proced" _
        , _
        "ure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=Host;Use Encryption for Data=False;Tag with column col" _
        , "lation when possible=False")
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("HealthBI")
        .Name = "HealthBI"
        .Description = ""
    End With
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=HealthBI;Data Source=CHH-BILive;Use Proced" _
        , _
        "ure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=Host;Use Encryption for Data=False;Tag with column col" _
        , "lation when possible=False"), Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("DECLARE @START AS Date DECLARE @END AS Date SET @START =" & "'" & StartDate & "'" & " SET @END =" & "'" & EndDate & "'" & " SELECT" _
        , _
        " mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_OP_REFERRAL ref INNER JOIN CDO_MPI mpi ON ref.CDO_MPI_UNIQUE_ID = mpi." _
        , _
        "UNIQUE_ID WHERE ref.EFFECTIVE_WAITING_START_DATE BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SE" _
        , _
        "LECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_APC_HOSPITAL_PROVIDER_SPELL adm INNER JOIN CDO_MPI mpi ON adm.CDO_" _
        , _
        "MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE adm.ADMISSION_METHOD_HOSPITAL_PROVIDER_SPELL NOT IN ('11','12','13') AND adm.ST" _
        , _
        "ART_DATE_HOSPITAL_PROVIDER_SPELL BETWEEN @START AND @END AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUM" _
        , _
        "BER, PERSON_BIRTH_DATE FROM CDO_OP_APPOINTMENT opa INNER JOIN CDO_MPI mpi ON opa.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID" _
        , _
        " WHERE (opa.APPOINTMENT_START_DATE BETWEEN @START AND @END OR opa.APPOINTMENT_BOOKED_DATE BETWEEN @START AND @END)" _
        , _
        " AND NHS_NUMBER IS NOT NULL UNION SELECT mpi.NHS_NUMBER, PERSON_BIRTH_DATE FROM CDO_A_AND_E_ATTENDANCE e" _
        , _
        "d INNER JOIN CDO_MPI mpi ON ed.CDO_MPI_UNIQUE_ID = mpi.UNIQUE_ID WHERE ed.ARRIVAL_DATE BETWEEN @START AND @END AND " _
        , "NHS_NUMBER IS NOT NULL GROUP BY mpi.NHS_NUMBER, PERSON_BIRTH_DATE")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "CHH_BILive_HealthBI"
        .Refresh BackgroundQuery:=False
    End With
    
End Sub

Sub Reformat_Date()
    
''FORMAT DATE OF BIRTH
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "m/d/yyyy"

End Sub
Sub OpenOVM()

''OPEN OVM TEMPLATE
    Sheets("NHS Numbers").Select
    Range("A2").Select
    
End Sub
Sub Copy_AND_Paste_alpha()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A2").Value <> "" Then

        ''COPY FIRST SET rows
            Range("A2:B" & SizeCheck).Select
                Selection.Copy
                
        ''PASTE FIRST SET rows
            Sheets("NHS Numbers").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select

        Else: GoTo Finished

End If

Finished:
  
End Sub
Sub Copy_AND_Paste_beta()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & (SizeCheck + 1)).Value <> "" Then

        ''COPY SECOND SET rows
            Range("A" & (SizeCheck + 1) & ":B" & (SizeCheck * 2) - 1).Select
                Selection.Copy
                
        ''PASTE SECOND SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished2

End If

Finished2:
    
End Sub
Sub Copy_AND_Paste_charlie()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 2) + 1)).Value <> "" Then

        ''COPY THIRD SET rows
            Range("A" & ((SizeCheck * 2) + 1) & ":B" & (SizeCheck * 3) - 1).Select
                Selection.Copy
                
        ''PASTE THIRD SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished3

End If

Finished3:
    
End Sub
Sub Copy_AND_Paste_delta()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 3) + 1)).Value <> "" Then

        ''COPY FOURTH SET rows
            Range("A" & ((SizeCheck * 3) + 1) & ":B" & (SizeCheck * 4) - 1).Select
                Selection.Copy
                
        ''PASTE FOURTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished4

End If

Finished4:
    
End Sub
Sub Copy_AND_Paste_echo()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 4) + 1)).Value <> "" Then

        ''COPY FIFTH SET rows
            Range("A" & ((SizeCheck * 4) + 1) & ":B" & (SizeCheck * 5) - 1).Select
                Selection.Copy
                
        ''PASTE FIFTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished5

End If

Finished5:
    
End Sub
Sub Copy_AND_Paste_foxtrot()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 5) + 1)).Value <> "" Then

        ''COPY SIXTH SET rows
            Range("A" & ((SizeCheck * 5) + 1) & ":B" & (SizeCheck * 6) - 1).Select
                Selection.Copy
                
        ''PASTE SIXTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished6

End If

Finished6:
    
End Sub

Sub Copy_AND_Paste_gamma()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 6) + 1)).Value <> "" Then

        ''COPY SEVENTH SET rows
            Range("A" & ((SizeCheck * 6) + 1) & ":B" & (SizeCheck * 7) - 1).Select
                Selection.Copy
                
        ''PASTE SEVENTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished7

End If

Finished7:
    
End Sub

Sub Copy_AND_Paste_hotel()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 7) + 1)).Value <> "" Then

        ''COPY SEVENTH SET rows
            Range("A" & ((SizeCheck * 7) + 1) & ":B" & (SizeCheck * 8) - 1).Select
                Selection.Copy
                
        ''PASTE SEVENTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished8

End If

Finished8:
    
End Sub

Sub Copy_AND_Paste_indigo()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 8) + 1)).Value <> "" Then

        ''COPY SEVENTH SET rows
            Range("A" & ((SizeCheck * 8) + 1) & ":B" & (SizeCheck * 9) - 1).Select
                Selection.Copy
                
        ''PASTE SEVENTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished9

End If

Finished9:
    
End Sub

Sub Copy_AND_Paste_juliet()

''GO BACK TO MESH_RAW
Sheets("Mesh_RAW").Select

    ''NEXT RANGE IS NOT BLANK
    If Range("A" & ((SizeCheck * 9) + 1)).Value <> "" Then

        ''COPY SEVENTH SET rows
            Range("A" & ((SizeCheck * 9) + 1) & ":B" & (SizeCheck * 10) - 1).Select
                Selection.Copy
                
        ''PASTE SEVENTH SET rows
            Sheets("NHS Numbers").Select
            Range("A2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            'ActiveWindow.SmallScroll Down:=-15
            'Sheets("OVM Request").Select
            
        Else: GoTo Finished10

End If

Finished10:
    
End Sub



