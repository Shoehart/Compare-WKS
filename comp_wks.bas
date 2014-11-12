Option Explicit

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
' Funkcja zwraca ilość zajmowanego ramu przez Excela
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
' Pobiera informacje do timerów, oraz ilość pamiecie zajmowanej przez Excela z Funkcji GetMemUsage
'================================================================================
Sub Pomiar_Start()
    QueryPerformanceFrequency uFreq
    QueryPerformanceCounter uStart
    memUseStart = GetMemUsage
End Sub

'================================================================================
' Sub Pomiar_Koniec
'
' Wrzuca pomiar czasu wykonywania funkcji od momentu zainicjowania przez Pomiar_Start
' do "Immediate", wraz z informacjami o ilości zajmowanego miejsca w ramie przez Excela
' na początku i na końcu wykonywania funkcji.
'================================================================================
Sub Pomiar_Koniec(nr As Long)
    QueryPerformanceCounter uEnd
    memUseEnd = GetMemUsage
    Debug.Print Format(Now, "hh") & ":" & Format(Now, "Nn") & " - Step #" & nr & ": " & Format((U64Dbl(uEnd) - U64Dbl(uStart)) / U64Dbl(uFreq), "0.000000"); " seconds elapsed." & " MemUsage (Start: " & Format(memUseStart / 1024, "0.00") & "MB , STOP: " & Format(memUseEnd / 1024, "0.00") & "MB. Difference: " & Format((memUseEnd - memUseStart) / 1024, "0.00") & "MB"
End Sub

'================================================================================
' Sub CompareWorksheets()
'
' Porównywanie dwóch otwartych arkuszy kalkulacyjnych, komórka po komórce.
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

'Wyłączenie wszystkich opcji spowalniających pracę Excela
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

'Wymiary poszczególnych arkuszy, czyli liczba kolumn i liczba wierszy
With wb1.Sheets(sA_WS).UsedRange
    lRow_1 = .Rows.Count
    lColumn_1 = .Columns.Count
End With
With wb2.Sheets(s2_WS).UsedRange
    lRow_2 = .Rows.Count
    lColumn_2 = .Columns.Count
End With

'wartości skrajne dla tablicy wynikowej/raportu
If lMaxR < lRow_2 Then lMaxR = lRow_2
If (lMaxC < lColumn_2) Or (lMaxC > 200) Then lMaxC = lColumn_2 'zabezpieczenie przed błędem w arkuszu (np. pokolorowany cała kolumna, co zwraca będne wartości dla UsedRange

'Zgranie danych z arkuszy do odpowiednich tablic
ReDim tempA(1 To lRow_1, 1 To 1)
ReDim tempB(1 To lRow_2, 1 To 1)
ReDim tempRaport(1 To lRow_2, 1 To lColumn_2)

'Licznik ilości różnic jaka wystąpiła między porównywanymi arkuszami
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

'Właściwa pętla porównująca ze sobą dwa arkusze. Ten mechanizm porównuje dzieląc badanych obszar
'na tyle fragmentów ile jest kolumn w danym zakresie.
If bCzyPrzeniescNaglowki = False Then
    i = 1
Else
    i = 2
    wb1.Activate
    Set rTempRng = wb1.Sheets(sA_WS).Range(Cells(1, 1), Cells(1, lMaxC))
    tempB = rTempRng.Value2
    For lTemp = 1 To lMaxC
        tempRaport(1, lTemp) = "'" & rTempRng(1, lTemp).Value
    Next lTemp
End If

lTemp = 0

For lColumn = 1 To lMaxC
    wb1.Activate
    Set rTempRng = wb1.Sheets(sA_WS).Range(Cells(1, lColumn), Cells(lRow_1, lColumn))
    tempA = rTempRng.Value2
    Set rTempRng = Nothing

    wb2.Activate
    Set rTempRng = wb2.Sheets(s2_WS).Range(Cells(1, lColumn), Cells(lRow_2, lColumn))
    tempB = rTempRng.Value2
    Set rTempRng = Nothing
    
    For lRow = i To lMaxR
        'Wrzucanie do raportu informacji o #Errorach w arkuszach (nie wychwytuje #REF, co wynika z kolejności wywoływania błędów przez Excela.
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

'Zazielenianie komórek w których nastapiła zmiana w ActiveWorkbook
If bGreenActiveWks Then
    wb1.Activate
    Set rTempRng = wb1.Sheets(sA_WS).Range(Cells(1, 1), Cells(lRow, lColumn))
    For lRow = i To UBound(tempRaport, 1)
        For lColumn = i To UBound(tempRaport, 2)
            If tempRaport(lRow, lColumn) <> vbNullString Then
                rTempRng(lRow, lColumn).Interior.ColorIndex = 43
            End If
        Next lColumn
    Next lRow
End If

'Zazielenianie komórek w których nastapiła zmiana w 2nd Workbook
If bGreen2ndWks Then
    wb2.Activate
    Set rTempRng = wb2.Sheets(s2_WS).Range(Cells(1, 1), Cells(lRow, lColumn))
    For lRow = i To UBound(tempRaport, 1)
        For lColumn = i To UBound(tempRaport, 2)
            If tempRaport(lRow, lColumn) <> vbNullString Then
                rTempRng(lRow, lColumn).Interior.ColorIndex = 43
            End If
        Next lColumn
    Next lRow
End If

'Zwolnienie pamięci ze zbędnych śmieci
Set rTempRng = Nothing
Unload frmCompWks
Erase tempA, tempB

Call Pomiar_Koniec(1)

Przygotowanie_Raportu nr_raportu:=lNrRaportu, czyNaglowki:=bCzyPrzeniescNaglowki, aRaport:=tempRaport, DiffCount:=lDiffCount
   
'Włączenie wszystkich funkcjonalności w Excelu spowrotem
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
' Sub do przygotowywania raportu z porównywania między sobą arkuszy. Na chwilę obecną
' dostępne są dwa. Uruchamia się to przekazując parametr liczbowy (od 1 do 2).
'================================================================================
Sub Przygotowanie_Raportu(ByVal nr_raportu As Long, ByVal czyNaglowki As Boolean, ByRef aRaport() As Variant, ByVal DiffCount As Long)

Dim tempRapKol() As Variant, tempRap() As Long
Dim lColumn As Long, lRow As Long, i As Long, xR As Long
Dim rptWB As Workbook
Dim tempS1 As String, lTemp As String
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
'Przygotowanie raportu z porównania - wersja #1
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
                'ActiveCell.FormulaR1C1 = "Jakiś link"
                'ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="124_porownywaczMoj.xlsm", TextToDisplay:="Jakiś link"
            End If
        Next lColumn
    Next lRow
    
    'wrzucanie poszczególnymi kolumnami do arkusza
    With ActiveSheet
        Range("A1").Value2 = "Adres komórki"
        Range("B1").Value2 = "MEMBER"
        Range("C1").Value2 = "Wartość z arkusza #1"
        Range("D1").Value2 = "Wartość z arkusza #2"
    End With
    Range(Cells(2, 1), Cells(UBound(tempRapKol, 1) + 1, 4)) = tempRapKol
        
   'Formatowanie całego skoroszytu przy wykorzystaniu Malarza Formatów
     With Range("A1")
        .FormatConditions.Add Type:=xlTextString, String:="Error", TextOperator:=xlContains
        .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
         With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
         End With
        .FormatConditions(1).StopIfTrue = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlHairline
        .Copy
    End With
    
    Range(Cells(1, 1), Cells(lTemp, 4)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Boldowanie nagłówków
    Range(Cells(1, 1), Cells(1, 4)).Select
    Selection.Font.Bold = True
    
    'rptWB.Save
    Set rptWB = Nothing
    lColumn = 5
    lRow = lTemp
'Przygotowanie raportu z porównania - wersja #2
Case 2
    Application.StatusBar = "Formatting the report (Style #" & nr_raportu & ")"
    'wrzucanie poszczególnymi kolumnami do arkusza
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
    
    With Range(Cells(xR, 1), Cells(UBound(aRaport, 1), 1)).FormatConditions
        .AddDatabar
    End With
        
   'Formatowanie całego skoroszytu przy wykorzystaniu Malarza Formatów
   'który kopiuje przygotowane formatowanie komórek z "A1" do całego zakresu.
    With Range("B2")
        '.Select
        .FormatConditions.Add Type:=xlExpression, Formula1:="=NOT(ISBLANK(B" & xR & "))=TRUE"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5296274
        End With
        .FormatConditions(1).StopIfTrue = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlHairline
        .Copy
    End With
        
    'Przeklejenie formatowania z "B2" do całego zakresu
    Range(Cells(xR, 2), Cells(UBound(aRaport, 1), UBound(aRaport, 2) + 1)).PasteSpecial _
        Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'od której kolumny i wiersza mają być schowane dane?
    lColumn = UBound(aRaport, 2) + 2
    lRow = UBound(aRaport, 1)
End Select

'Ustawienie autoSzerokości kolumn w odniesieniu do ich zawartości, lecz AutoFit nie szerszy niż 350
For Each rTempCell In ActiveSheet.UsedRange.Rows(1).Cells
    If rTempCell.Column <> 1 Then
        With rTempCell.EntireColumn
            .AutoFit
            If .ColumnWidth > 50 Then
               .ColumnWidth = 50
            End If
        End With
    Else
        rTempCell.EntireColumn.ColumnWidth = 15
    End If
Next rTempCell
    
With Range(Cells(1, 1), Cells(lRow, lColumn))
    .AutoFilter
End With

'Podliczenie zmian w ostatnim wierszu, dla każdej kolumny
With Range(Cells(lRow + 1, 2), Cells(lRow + 1, lColumn + 1))
    .FormulaR1C1 = "=COUNTA(R" & xR & "C:R[-1]C)"
    .Interior.ColorIndex = 15
End With
    
'Chowanie zbędnych wierszy
Range(Rows(lRow + 2), Rows(lRow + 2).End(xlDown)).EntireRow.Hidden = True
'Chowanie zbędnych kolumn
Range(Columns(lColumn), Columns(lColumn).End(xlToRight)).EntireColumn.Hidden = True

'rptWB.Save
Set rptWB = Nothing

'End of timing
Call Pomiar_Koniec(2)

End Sub

'================================================================================
' Sub Update_Progress(ByVal Wartosc As Long)
'
' Uaktualnia progressBar o podaną wartość
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
' Głowna funkcja startująca formularz do porównana danych. Zbiera informacje o wybranych
' workbookach, i opcje zielenienia różnica w workbokach wajściowych.
'================================================================================
Sub comp_wks()
    Dim WSNames() As String
    Dim WBNames() As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    Dim N As Byte, M As Byte
    Dim CompareAll As Boolean
    
    Dim sActiveWB As String
    Dim sActiveWS As String
    Dim s2ndWB As String
    Dim s2ndWS As String
    Dim lNrRaportu As Long
    
    Dim lorkszit1, lorkszit2 As Worksheet
    Dim aktywny As Workbook
    Dim x, y, z, identical As Integer
    identical = 0
        
    bGreenActiveWks = False
    bGreen2ndWks = False
    CompareAll = False
        
    If GetWksCount = 0 Then
      MsgBox "Please open at least 2 Workbooks!", vbOKOnly
      Exit Sub
    End If
    
    M = 0
    i = 0
    ReDim WBNames(0 To Workbooks.Count - 1)
    For Each WB In Workbooks
        If (WB.Name <> ActiveWorkbook.Name) And (WB.Name <> "PERSONAL.XLSB") Then
            ReDim WSNames(0 To WB.Worksheets.Count)
            WBNames(M) = WB.Name
            M = M + 1
        End If
    Next WB
    
    'Wgranie danych do formularz obsługujacego porównywanie arkuszy excelowych
    Load frmCompWks
    With frmCompWks
      .cboActiveWB.Clear
      .cboActiveWks.Clear
      .cbo2ndWks.Clear
      .cbo2ndWB.Clear
      .cbo2ndWB.ListIndex = -1
      .cmdOK.Enabled = False
      .Height = 312
    
    'Uzupełnienie pól dla ActiveWorkbook / oraz Sheets z tego AW
    .cboActiveWB.AddItem ActiveWorkbook.Name, -1
    For Each WS In Worksheets
      .cboActiveWks.AddItem WS.Name, i
      i = i + 1
    Next
      
    'Uzupełnienie ListBoxa o pozostałe WB
    For i = 0 To UBound(WBNames) - 1
        .cbo2ndWB.AddItem WBNames(i), i
    Next i
    
    'Uzupełnienie o Sheets z pozostałych WB.
    For i = 1 To 2
        .cboChooseRaport.AddItem i
    Next i
      
    Erase WSNames(), WBNames()
    
      .cboActiveWB.ListIndex = 0
      .cboActiveWks.ListIndex = 0
      .cbo2ndWks.ListIndex = -1
      .cboChooseRaport.ListIndex = 0
    
      'display it
      .Show
      
      '.Tag True oznacza, że w formularzu został naciśnięty przycisk OK i że mają być wykonane obliczenia.
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
      
End With
'    If CompareAll = True Then
'        Set rptWBAll = Workbooks.Add
'        Application.DisplayAlerts = False
'        While Worksheets.Count > 1
'            Worksheets(2).Delete
'        Wend
'        Application.DisplayAlerts = True
'        rptWBAll.Worksheets(1).Name = "Error Log CmpWs"
'        rptWBAll.Worksheets(1).Range("a1") = "Active Workbook"
'        rptWBAll.Worksheets(1).Range("b1") = "Compared Workbook"
'        rptWBAll.Worksheets(1).Range("c1") = "Diff Count"
'        rptWBAll.Worksheets(1).Range("a1:c1").Font.Bold = True
'        y = 0
'        For Each lorkszit1 In aktywny.Worksheets
'            x = 0
'            For Each lorkszit2 In Workbooks(s2ndWB).Worksheets
'                If lorkszit1.Name = lorkszit2.Name Then
'                    On Error Resume Next
'                    CompareWorksheetsAll aktywny.Worksheets(lorkszit1.Name), _
'                        Workbooks(s2ndWB).Worksheets(lorkszit2.Name)
'                    On Error GoTo 0
'                    rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 0) = lorkszit1.Name
'                    rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 2) = lDiffCount
'                    rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 1) = lorkszit2.Name
'                    If lDiffCount <> 0 Then
'                        rptWBAll.Worksheets(1).Range("a1:c1").Offset(y + 1, 0).Interior.ColorIndex = 38
'                        identical = identical + 1
'                    End If
'                    x = x + 1
'                End If
'
'            Next
'            If x = 0 Then
'                rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 0) = lorkszit1.Name
'                rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 1) = "N/A"
'                rptWBAll.Worksheets(1).Range("a1:c1").Offset(y + 1, 0).Interior.ColorIndex = 40
'                identical = identical + 1
'            End If
'            y = y + 1
'        Next
'
'        For Each lorkszit2 In Workbooks(s2ndWB).Worksheets
'            z = 0
'            For Each lorkszit1 In aktywny.Worksheets
'                If lorkszit1.Name = lorkszit2.Name Then
'                    z = z + 1
'                End If
'            Next
'            If z = 0 Then
'                rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 0) = "N/A"
'                rptWBAll.Worksheets(1).Range("a1").Offset(y + 1, 1) = lorkszit2.Name
'                rptWBAll.Worksheets(1).Range("a1:c1").Offset(y + 1, 0).Interior.ColorIndex = 40
'                identical = identical + 1
'                y = y + 1
'            End If
'        Next
'
'        rptWBAll.Worksheets(1).Activate
'        rptWBAll.Worksheets(1).Columns("a:c").AutoFit
'        rptWBAll.Saved = True
'        If identical = 0 Then
'            MsgBox "Workbooks are identical"
'        End If
'    Else
        'On Error GoTo ErrHandler
        CompareWorksheets sA_WB:=sActiveWB, sA_WS:=sActiveWS, s2_WB:=s2ndWB, s2_WS:=s2ndWS
'    End If
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
' Zwraca liczbowo ilość otwartych Workbooków, bez PERSONAL.XLSB
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
