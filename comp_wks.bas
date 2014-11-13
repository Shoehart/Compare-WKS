Attribute VB_Name = "Comp_wks"
Option Explicit

'Need to be done:
'1. Make it EX2013 compatibile, so 64bit
'2. Code cleaning
'3. Translation all comments to english
'4. Compare many worksheets
'5. Check for #REF
'6. AutoFilter without last row in raport - DONE!
'7. Hiperlinki in raports to apropriate worksheets

Private Type UINT64
    LowPart As Long
    HighPart As Long
End Type
Private Const BSHIFT_32 = 4294967296# ' 2 ^ 32

'#If VBA7 Then
'    #If Win64 Then
'        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As LongPtr
'        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As LongPtr
'        Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
'    #End If
'#Else
        Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As Long
        Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As Long
        Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'#End If

Public bGreenActiveWks As Boolean
Public bGreen2ndWks As Boolean

Dim memUseStart As Long
Dim memUseEnd As Long
    
'zmienne do timera
Dim uStart As UINT64
Dim uEnd As UINT64
Dim uFreq As UINT64
Dim dblElapsed As Double

Dim rptWBAll As Workbook

Private Function U64Dbl(U64 As UINT64) As Double
    Dim lDbl As Double, hDbl As Double
    lDbl = U64.LowPart
    hDbl = U64.HighPart
    If lDbl < 0 Then lDbl = lDbl + BSHIFT_32
    If hDbl < 0 Then hDbl = hDbl + BSHIFT_32
    U64Dbl = lDbl + BSHIFT_32 * hDbl
End Function

'================================================================================
' Sub GetMemUsage
'
' Get amount of RAM memory Excel eat while running
'================================================================================
Function GetMemUsage()
Dim objSWbemServices As Object
  Set objSWbemServices = GetObject("winmgmts:")
      GetMemUsage = objSWbemServices.Get( _
      "Win32_Process.Handle='" & _
      GetCurrentProcessId & "'").WorkingSetSize / 1024
  Set objSWbemServices = Nothing
End Function

'================================================================================
' Sub Pomiar_Start
'
' Sub used for timers. Start of counting, and also uses func GetMemUsage
'================================================================================
Sub Pomiar_Start()
    QueryPerformanceFrequency uFreq
    QueryPerformanceCounter uStart
    memUseStart = GetMemUsage
End Sub

'================================================================================
' Sub Pomiar_Koniec
'
' Counts time elapsed from the start of "Pomiar_start". Prints it in Immediate,
' also prints amount of RAM used by Excel on the begining and end of using of measured function.
'================================================================================
Sub Pomiar_Koniec(nr As Long)
    QueryPerformanceCounter uEnd
    memUseEnd = GetMemUsage
    Debug.Print Format(Now, "hh") & ":" & Format(Now, "Nn") & " - Step #" & nr & ": " & Format((U64Dbl(uEnd) - U64Dbl(uStart)) / U64Dbl(uFreq), "0.000000"); " seconds elapsed." & " MemUsage (Start: " & Format(memUseStart / 1024, "0.00") & "MB , STOP: " & Format(memUseEnd / 1024, "0.00") & "MB. Difference: " & Format((memUseEnd - memUseStart) / 1024, "0.00") & "MB"
End Sub

'================================================================================
' Sub CompareWorksheets()
'
' Main sub comparing two given worksheets. Cell by cell. Sub has a little bit of error
' checking f.e. #NAME. #REF still needs to be done.
'================================================================================
Sub CompareWorksheets(ByVal sA_WB As String, _
                      ByVal sA_WS As String, _
                      ByVal s2_WB As String, _
                      ByVal s2_WS As String)

Dim lRow As Long, lColumn As Long, _
    lRow_1 As Long, lColumn_1 As Long, _
    lRow_2 As Long, lColumn_2 As Long

Dim lMaxR As Long, lMaxC As Long
Dim lDiffCount As Long
Dim lDzielnik As Long, lCount As Double
Dim lNrRaportu As Long

'All temp variables
Dim tempS1 As String, tempS2 As String, sTemp As String
Dim tempA() As Variant, tempB() As Variant, tempRaport() As Variant, tempRapKol() As Variant
Dim lTemp As Long
Dim i As Long
Dim rTempRng As Range

Dim rptWB As Workbook
Dim wb1 As Workbook, wb2 As Workbook

Dim bCzyZrobicRaport As Boolean
Dim bCzyPrzeniescNaglowki As Boolean

'Turn off all Excel variable slowing down program
With Application
    .EnableEvents = False
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .StatusBar = False
    .ErrorCheckingOptions.NumberAsText = False
    .SheetsInNewWorkbook = 1
End With

Pomiar_Start
    
Set wb1 = Workbooks(sA_WB)
Set wb2 = Workbooks(s2_WB)

'Check size of UsedRange of compared worksheets
With wb1.Sheets(sA_WS).UsedRange
    lRow_1 = .Rows.Count
    lColumn_1 = .Columns.Count
End With
With wb2.Sheets(s2_WS).UsedRange
    lRow_2 = .Rows.Count
    lColumn_2 = .Columns.Count
End With

'Max values for number of rows and columns for two compared worksheets. It is important
'for raport generation.
If lMaxR < lRow_2 Then lMaxR = lRow_2
If (lMaxC < lColumn_2) Or (lMaxC > 200) Then lMaxC = lColumn_2 'zabezpieczenie przed b³êdem w arkuszu (np. pokolorowany ca³a kolumna, co zwraca bêdne wartoœci dla UsedRange

'Copy all data from worksheets to array (but, column by column - to evade "Out of Memory" error)
ReDim tempA(1 To lRow_1, 1 To 1)
ReDim tempB(1 To lRow_2, 1 To 1)
ReDim tempRaport(1 To lRow_2, 1 To lColumn_2)

'Variable for counting a number of differance
lDiffCount = 0

lCount = lMaxC * lMaxR
lDzielnik = lCount \ 100

With frmCompWks
    .Height = 334
    .Show vbModeless
    .ProgressBar.Visible = True
    .ProgressBar.Enabled = True
    .ProgressBar.Min = 0
    .ProgressBar.Max = lCount
    lNrRaportu = .cboChooseRaport.Value
    bCzyPrzeniescNaglowki = .cboHeader.Value
    bGreenActiveWks = .CheckBox1.Value
    bGreen2ndWks = .CheckBox2.Value
End With

'Main code for comparing data. In this aproach data are divided in to columns. As
'many as there is columns in biggest compared worksheets.

'Checkup for "Headers". If Yes then first row isn't check, and raport has headers from raw data.
If bCzyPrzeniescNaglowki = False Then
    i = 1
Else
    i = 2
    'wb1.Activate
    wb1.Sheets(sA_WS).Activate
    Set rTempRng = wb1.Sheets(sA_WS).Range(Cells(1, 1), Cells(1, lMaxC))
    tempB = rTempRng.Value2
    For lTemp = 1 To lMaxC
        tempRaport(1, lTemp) = "'" & rTempRng(1, lTemp).Value2
    Next lTemp
End If

lTemp = 0

'Check: column by column
For lColumn = 1 To lMaxC
    'wb1.Activate
    wb1.Sheets(sA_WS).Activate
    Set rTempRng = wb1.Sheets(sA_WS).Range(Cells(1, lColumn), Cells(lRow_1, lColumn))
    tempA = rTempRng.Value2
    Set rTempRng = Nothing

    'wb2.Activate
    wb2.Sheets(s2_WS).Activate
    Set rTempRng = wb2.Sheets(s2_WS).Range(Cells(1, lColumn), Cells(lRow_2, lColumn))
    tempB = rTempRng.Value2
    Set rTempRng = Nothing
    
    For lRow = i To lMaxR
        'Input msg on Error in a worksheets data. Doesn't catch #REF.
        'Main reason is the order of error cheecking in Excel. Needs to be done.
        If IsError(tempA(lRow, 1)) Then
            tempA(lRow, 1) = "Error"
        End If
        If IsError(tempB(lRow, 1)) Then
            tempB(lRow, 1) = "Error"
        End If
        
        tempS1 = tempA(lRow, 1)
        tempS2 = tempB(lRow, 1)
        
            'The most important part of macro - the TEST!
            If tempS1 <> tempS2 Then
                lDiffCount = lDiffCount + 1
                tempRaport(lRow, lColumn) = "'" & tempS1 & " <> " & tempS2
            End If
            
            'Pushing date to ProgressBar object
            If lRow Mod lDzielnik = 0 Then
                lTemp = lTemp + lDzielnik
                Update_Progress Wartosc:=lTemp
            End If
    Next lRow
Next lColumn

'Setting green background color in cells in compared worksheets. This part is responsible for ActiveWorkbook
If bGreenActiveWks Then
    'wb1.Activate
    wb1.Sheets(sA_WS).Activate
    Set rTempRng = wb1.Sheets(sA_WS).Range(Cells(1, 1), Cells(lRow, lColumn))
    For lRow = i To UBound(tempRaport, 1)
        For lColumn = i To UBound(tempRaport, 2)
            If tempRaport(lRow, lColumn) <> vbNullString Then
                rTempRng(lRow, lColumn).Interior.ColorIndex = 43
            End If
        Next lColumn
    Next lRow
End If

'Setting green background color in cells in compared worksheets. This part is responsible for 2nd WorkBook
If bGreen2ndWks Then
    'wb2.Activate
    wb2.Sheets(s2_WS).Activate
    Set rTempRng = wb2.Sheets(s2_WS).Range(Cells(1, 1), Cells(lRow, lColumn))
    For lRow = i To UBound(tempRaport, 1)
        For lColumn = i To UBound(tempRaport, 2)
            If tempRaport(lRow, lColumn) <> vbNullString Then
                rTempRng(lRow, lColumn).Interior.ColorIndex = 43
            End If
        Next lColumn
    Next lRow
End If

'GarbageCollector ;)
Set rTempRng = Nothing
Unload frmCompWks
Erase tempA, tempB

Call Pomiar_Koniec(1)

Przygotowanie_Raportu nr_raportu:=lNrRaportu, czyNaglowki:=bCzyPrzeniescNaglowki, aRaport:=tempRaport, DiffCount:=lDiffCount
   
'Turn on all Excel variable, turned off at the beggining of the script
With Application
    .EnableEvents = True
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .StatusBar = False
    .ErrorCheckingOptions.NumberAsText = True
    .SheetsInNewWorkbook = 3
End With
End Sub

'================================================================================
' Przygotowanie_Raportu(ByVal nr_raportu As Long, ByRef aRaport() As Variant, ByVal DiffCount As Long)
'
' Sub do przygotowywania raportu z porównywania miêdzy sob¹ arkuszy. Na chwilê obecn¹
' dostêpne s¹ dwa. Uruchamia siê to przekazuj¹c parametr liczbowy (od 1 do 2).
'================================================================================
Sub Przygotowanie_Raportu(ByVal nr_raportu As Long, _
                          ByVal czyNaglowki As Boolean, _
                          ByRef aRaport() As Variant, _
                          ByVal DiffCount As Long)

Dim lColumn As Long, lRow As Long, i As Long, xR As Long

Dim rptWB As Workbook

'temp variables
Dim tempRapKol() As Variant, tempRap() As Long, lTemp As Long
Dim tempS1 As String, sTemp As String
Dim rTempCell As Range

Pomiar_Start

If DiffCount > 0 Then
    Set rptWB = Workbooks.Add
End If

If czyNaglowki Then
    xR = 2
Else
    xR = 1
End If

Select Case nr_raportu
'Prepare raport #1
Case 1
    Application.StatusBar = "Formatting the report (Style #" & nr_raportu & ")"
    ReDim tempRapKol(1 To UBound(aRaport, 1), 1 To 4)
    lTemp = 1
    For lRow = xR To UBound(aRaport, 1)
        For lColumn = 1 To UBound(aRaport, 2)
            If aRaport(lRow, lColumn) <> vbNullString Then
                tempS1 = Application.ConvertFormula("R" & lRow & "C" & lColumn, FromReferenceStyle:=xlR1C1, ToReferenceStyle:=xlA1)
                tempRapKol(lTemp, 1) = "Record: " & lRow & " " & tempS1 '& " (Wiersz: " & lRow & ", Kolumna: " & lColumn & ")"
                tempRapKol(lTemp, 2) = aRaport(1, lColumn)
                tempRapKol(lTemp, 3) = "'" & Mid(aRaport(lRow, lColumn), 2, InStr(aRaport(lRow, lColumn), "<>") - 3)
                tempRapKol(lTemp, 4) = "'" & Right(aRaport(lRow, lColumn), Len(aRaport(lRow, lColumn)) - InStr(aRaport(lRow, lColumn), "<>") - 2)
                lTemp = lTemp + 1
                'ActiveCell.FormulaR1C1 = "Jakiœ link"
                'ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="124_porownywaczMoj.xlsm", TextToDisplay:="Jakiœ link"
            End If
        Next lColumn
    Next lRow
    
    'Column by column
    With ActiveSheet
        Range("A1").Value2 = "Adres komórki"
        Range("B1").Value2 = "MEMBER"
        Range("C1").Value2 = "Wartoœæ z arkusza #1"
        Range("D1").Value2 = "Wartoœæ z arkusza #2"
    End With
    Range(Cells(2, 1), Cells(UBound(tempRapKol, 1) + 1, 4)) = tempRapKol
        
    'Formating titles: Bold
    Range(Cells(1, 1), Cells(1, 4)).Font.Bold = True
    
    'rptWB.Save
    Set rptWB = Nothing
    lColumn = 5
    lRow = lTemp
    
'Prepare raport #2
Case 2
    Application.StatusBar = "Formatting the report (Style #" & nr_raportu & ")"
    'Column by column in to sheet
    ReDim tempRapKol(1 To UBound(aRaport, 1), 1 To 1)
    ReDim tempRap(1 To UBound(aRaport, 1), 1 To 1)
    
    For lColumn = 1 To UBound(aRaport, 2)
        For lRow = 1 To UBound(aRaport, 1)
            tempRapKol(lRow, 1) = aRaport(lRow, lColumn)
            If aRaport(lRow, lColumn) <> vbNullString Then
                tempRap(lRow, 1) = tempRap(lRow, 1) + 1
            End If
        Next lRow
        Range(Cells(1, lColumn + 1), Cells(UBound(aRaport, 1), lColumn + 1)) = tempRapKol
    Next lColumn
    
    Range(Cells(1, 1), Cells(UBound(aRaport, 1), 1)) = tempRap
    
    'size of data in raport
    lColumn = UBound(aRaport, 2) + 2
    lRow = UBound(aRaport, 1)
End Select

'Formating raport with "Table"
With ActiveSheet
    'Checking if future Table will have Headers or not and creation of Table with few variables
    If xR = 1 Then
        .ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, lColumn - 1)), , xlNo).Name = "Raport_" & nr_raportu & Chr(34)
    Else
        .ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, lColumn - 1)), , xlYes).Name = "Raport_" & nr_raportu & Chr(34)
    End If
    With .ListObjects("Raport_" & nr_raportu & Chr(34))
        .ShowTableStyleRowStripes = False
        .ShowTotals = True
        .ShowAutoFilter = True
        .TableStyle = "TableStyleMedium2"
        For i = 2 To .ListColumns.Count
          .ListColumns(i).TotalsCalculation = xlTotalsCalculationCount
        Next i
    End With
     
If nr_raportu <> 1 Then
    'Add DataBar in added column
    Range(Cells(xR, 1), Cells(lRow, 1)).FormatConditions.AddDatabar
    'Add Conditional Formating
    With Range(Cells(xR, 2), Cells(lRow, lColumn))
        .FormatConditions.Add Type:=xlExpression, Formula1:="=NOT(ISBLANK(B2))=TRUE"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5296274
        End With
        .FormatConditions(1).StopIfTrue = False
        '.Copy
    End With
End If
    
i = 0
'AutoFit columns. Not wider than 50 point. Empty columns set to 2
For Each rTempCell In ActiveSheet.UsedRange.Rows(1).Cells
i = i + 1
    If nr_raportu <> 1 Then
        If rTempCell.Column <> 1 Then
            With rTempCell.EntireColumn
                If Cells(lRow + 1, i) = 0 Then
                    .ColumnWidth = 2
                Else
                    .AutoFit
                    Cells(1, i).Interior.ColorIndex = 50
                    If .ColumnWidth > 50 Then
                       .ColumnWidth = 50
                    End If
                End If
            End With
        Else
            rTempCell.EntireColumn.ColumnWidth = 8
        End If
    Else
        With rTempCell.EntireColumn
            .AutoFit
            If .ColumnWidth > 60 Then
               .ColumnWidth = 60
            End If
        End With
    End If
Next rTempCell
End With

'Hide unused rows
Range(Rows(lRow + 2), Rows(lRow + 2).End(xlDown)).EntireRow.Hidden = True
'Hide unused columns
Range(Columns(lColumn), Columns(lColumn).End(xlToRight)).EntireColumn.Hidden = True

'rptWB.Save
Set rptWB = Nothing

'End of timing
Call Pomiar_Koniec(2)

End Sub

'================================================================================
' Sub Update_Progress(ByVal Wartosc As Long)
'
' Uaktualnia progressBar o podan¹ wartoœæ
'================================================================================
Sub Update_Progress(ByVal Wartosc As Long)
    With frmCompWks.ProgressBar
        If Wartosc < .Max Then
            .Refresh
            .Value = Wartosc
        End If
    End With
    DoEvents
End Sub
'================================================================================
' Sub CompareWorksheetsAll()
'
' Porównywanie wszystkich arkuszy w dwóch otwartych Workbookach, komórka po komórce.
' Wraz z wygenerowaniem raportu.
'================================================================================
Sub CompareWorksheetsAll(ws1 As Worksheet, ws2 As Worksheet)

Dim r As Long, c As Integer
Dim lRow_1 As Long, lRow_2 As Long, lColumn_1 As Integer, lColumn_2 As Integer
Dim lMaxR As Long, lMaxC As Integer, tempS1 As String, tempS2 As String
Dim rptWS As Worksheet, lDiffCount As Long

    Application.ScreenUpdating = False
    Application.StatusBar = "Creating the report..."

    Application.DisplayAlerts = False
    Set rptWS = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    rptWS.Name = ws1.Name
    
    Application.DisplayAlerts = True
    With ws1.UsedRange
        lRow_1 = .Rows.Count
        lColumn_1 = .Columns.Count
    End With
    With ws2.UsedRange
        lRow_2 = .Rows.Count
        lColumn_2 = .Columns.Count
    End With
    lMaxR = lRow_1
    lMaxC = lColumn_1
    If lMaxR < lRow_2 Then lMaxR = lRow_2
    If lMaxC < lColumn_2 Then lMaxC = lColumn_2
    lDiffCount = 0
    For c = 1 To lMaxC
        Application.StatusBar = "Comparing cells " & Format(c / lMaxC, "0 %") & "..."
        For r = 1 To lMaxR
            tempS1 = ""
            tempS2 = ""
            On Error Resume Next
            tempS1 = ws1.Cells(r, c).Value
            tempS2 = ws2.Cells(r, c).Value
            On Error GoTo 0
            If tempS1 <> tempS2 Then
                lDiffCount = lDiffCount + 1
                Cells(r, c).Formula = "'" & tempS1 & " <> " & tempS2
                Cells(r, c).Interior.ColorIndex = 44
                If bGreenActiveWks = True Then
                    ws1.Cells(r, c).Interior.ColorIndex = 43
                End If
                If bGreen2ndWks = True Then
                    ws2.Cells(r, c).Interior.ColorIndex = 43
                End If
            End If
        Next r
    Next c
    Application.StatusBar = "Formatting the report..."
    With Range(Cells(1, 1), Cells(lMaxR, lMaxC))
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        On Error GoTo 0
    End With
    Range(Cells(lMaxR + 1, 1), Cells(lMaxR + 1, lMaxC)).FormulaR1C1 = "=COUNTA(R1C:R[-1]C)"
    Cells.Select
    Cells.EntireColumn.AutoFit
    Rows(lMaxR + 1).Interior.ColorIndex = 15
    Rows(lMaxR + 2).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Hidden = True
'    rptWB.Saved = True
    Application.DisplayAlerts = False
    If lDiffCount = 0 Then
        rptWS.Delete
    End If
    Application.DisplayAlerts = True
    Set rptWS = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

'================================================================================
' Sub comp_wks()
'
' G³owna funkcja startuj¹ca formularz do porównana danych. Zbiera informacje o wybranych
' workbookach, i opcje zielenienia ró¿nica w workbokach wajœciowych.
'================================================================================
Sub comp_wks()
    Dim WSNames() As String
    Dim WBNames() As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    Dim N As Byte
    Dim bCompareAll As Boolean
    
    Dim sActiveWB As String
    Dim sActiveWS As String
    Dim s2ndWB As String
    Dim s2ndWS As String
    Dim lNrRaportu As Long
    Dim lDiffCount As Long
    
    Dim s1_WS, s2_WS As Worksheet
    Dim aktywny As Workbook
    Dim x, y, z, identical As Long
    
    identical = 0
        
    bGreenActiveWks = False
    bGreen2ndWks = False
    bCompareAll = False
        
    If GetWksCount = 0 Then
      MsgBox "Please open at least 2 Workbooks!", vbOKOnly
      Exit Sub
    End If
    
    i = 0
    ReDim WBNames(0 To Workbooks.Count - 1)
    For Each WB In Workbooks
        If (WB.Name <> ActiveWorkbook.Name) And (WB.Name <> "PERSONAL.XLSB") Then
            ReDim WSNames(0 To WB.Worksheets.Count)
            WBNames(i) = WB.Name
            i = i + 1
        End If
    Next WB
    
    'Settings for macro, set in form.
    Load frmCompWks
    With frmCompWks
      .cboActiveWB.Clear
      .cboActiveWks.Clear
      .cbo2ndWks.Clear
      .cbo2ndWB.Clear
      .cboChooseRaport.Clear
      .cmdOK.Enabled = True
      .Height = 312
    
    i = 0
    'Filling field with ActiveWorkbook name / and Sheets names in next field
    .cboActiveWB.AddItem ActiveWorkbook.Name, -1
    For Each WS In Worksheets
      .cboActiveWks.AddItem WS.Name, i
      i = i + 1
    Next
      
    'Filling ListBox with the name of the rest WB
    For i = 0 To UBound(WBNames) - 1
        .cbo2ndWB.AddItem WBNames(i), i
    Next i
    
    i = 0
    For Each WS In Workbooks(.cbo2ndWB.List(0)).Worksheets
        .cbo2ndWks.AddItem WS.Name, i
        i = i + 1
    Next

    'Fill rest of Sheets from the rest of WB
    For i = 1 To 2
        .cboChooseRaport.AddItem i
    Next i
      
    Erase WSNames(), WBNames()
    
      .cboActiveWB.ListIndex = 0
      .cboActiveWks.ListIndex = 0
      .cbo2ndWks.ListIndex = 0 'set to 0 for testing. Need to be changed to "-1"
      .cboChooseRaport.ListIndex = 0
    
      'display it
      .Show
      
      '.Tag True oznacza, ¿e w formularzu zosta³ naciœniêty przycisk OK i ¿e maj¹ byæ wykonane obliczenia.
      If .Tag = "True" Then
          Unload frmCompWks
          Exit Sub
      End If
      
      sActiveWB = .cboActiveWB.Value
      For i = 0 To .cbo2ndWB.ListCount - 1
          If .cbo2ndWB.Selected(i) Then
              s2ndWB = .cbo2ndWB.List(i)
          End If
      Next i
      
      sActiveWS = .cboActiveWks.Value
      s2ndWS = .cbo2ndWks.Value
      bCompareAll = .cboAllTabs
      
End With
    
'Do poprawienia!
If bCompareAll = False Then
    'On Error GoTo ErrHandler
    CompareWorksheets sA_WB:=sActiveWB, sA_WS:=sActiveWS, s2_WB:=s2ndWB, s2_WS:=s2ndWS
Else
    With Application
        .SheetsInNewWorkbook = 1
        .DisplayAlerts = False
    End With
    
    Set rptWBAll = Workbooks.Add
    Set aktywny = Workbooks(sActiveWB)
    
    Application.DisplayAlerts = True
    
    With rptWBAll.Worksheets(1)
        .Name = "Error Log CmpWs"
        .Range("A1") = "Active Workbook"
        .Range("B1") = "Compared Workbook"
        .Range("C1") = "Diff Count"
        .Range("A1:C1").Font.Bold = True
    End With
    
    y = 0
    For Each s1_WS In Workbooks(sActiveWB).Worksheets
        x = 0
        For Each s2_WS In Workbooks(s2ndWB).Worksheets
            If s1_WS.Name = s2_WS.Name Then
                On Error Resume Next
                
                'CompareWorksheetsAll aktywny.Worksheets(lorkszit1.Name), _
                    Workbooks(s2ndWB).Worksheets(lorkszit2.Name)
                CompareWorksheets sActiveWB, s1_WS.Name, s2ndWB, s2_WS.Name
                    
                On Error GoTo 0
                With rptWBAll.Worksheets(1)
                    .Range("a1").Offset(y + 1, 0) = s1_WS.Name
                    .Range("a1").Offset(y + 1, 2) = 1 'lDiffCount - poprawiæ przekazywanie iloœci ró¿nic
                    .Range("a1").Offset(y + 1, 1) = s2_WS.Name
                    If lDiffCount <> 0 Then
                        .Range("a1:c1").Offset(y + 1, 0).Interior.ColorIndex = 38
                        identical = identical + 1
                    End If
                    x = x + 1
                End With
            End If
        Next
        If x = 0 Then
        With rptWBAll.Worksheets(1)
            .Range("a1").Offset(y + 1, 0) = s1_WS.Name
            .Range("a1").Offset(y + 1, 1) = "N/A"
            .Range("a1:c1").Offset(y + 1, 0).Interior.ColorIndex = 40
            identical = identical + 1
        End With
        End If
        y = y + 1
    Next

    For Each s2_WS In Workbooks(s2ndWB).Worksheets
        z = 0
        For Each s1_WS In Workbooks(sActiveWB).Worksheets
            If s1_WS.Name = s2_WS.Name Then
                z = z + 1
            End If
        Next
        If z = 0 Then
        With rptWBAll.Worksheets(1)
            .Range("A1").Offset(y + 1, 0) = "N/A"
            .Range("A1").Offset(y + 1, 1) = s2_WS.Name
            .Range("A1:C1").Offset(y + 1, 0).Interior.ColorIndex = 40
            identical = identical + 1
            y = y + 1
        End With
        End If
    Next
    
    With rptWBAll.Worksheets(1)
        .Activate
        .Columns("A:C").AutoFit
    End With
    rptWBAll.Saved = True
    If identical = 0 Then
        MsgBox "Workbooks are identical"
    End If
End If
'Exit Sub
'ErrHandler:
'    Select Case Err.Number
'        Case 9
'            Call MsgBox(Err.Description & vbCrLf & vbCrLf & "No such worksheet", vbCritical + vbOKOnly, "Error")
'            Err.Clear
'            Resume Next
'        Case Else
'        ' All outstanding errors
'            MsgBox Err.Number & ": " & Error.Description
'            Err.Clear
'            Resume Next
'    End Select
End Sub

'================================================================================
' Function GetWksCount() As Long
'
' Returns number of opened WB without PERSONAL.XLSB (this wont work in EX2013)
'================================================================================
Private Function GetWksCount() As Long
    Dim WB As Workbook
    Dim N As Long

    N = 0
    For Each WB In Workbooks
      If WB.Name <> "PERSONAL.XLSB" Then
          N = N + 1
      End If
    Next WB
    GetWksCount = N - 1
End Function
