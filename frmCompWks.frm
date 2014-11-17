VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompWks 
   Caption         =   "Choose file and tab to compare"
   ClientHeight    =   6255
   ClientLeft      =   1050
   ClientTop       =   2370
   ClientWidth     =   4530
   OleObjectBlob   =   "frmCompWks.frx":0000
End
Attribute VB_Name = "frmCompWks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo1stWB_Click()
Dim i As Long
Dim WS As Worksheet
Dim WB As Workbook

'Wype³nienie ComboBox'a cbo1stWB nazwami workbookow.
With frmCompWks
    .cbo1stWks.Clear
    
    i = 0
    For Each WS In Workbooks(.cbo1stWB.Value).Worksheets
        .cbo1stWks.AddItem WS.Name, i
        i = i + 1
    Next
    
    .cbo2ndWB.Clear
    
    i = 0
    For Each WB In Workbooks
        If (WB.Name <> "PERSONAL.XLSB") And (WB.Name <> .cbo1stWB.Value) Then
            .cbo2ndWB.AddItem WB.Name, i
            i = i + 1
        End If
    Next
    .cbo1stWks.ListIndex = 0
End With

SprawdzCombosy

End Sub

Private Sub cbo2ndWB_Click()
Dim i As Long
Dim WS As Worksheet
Dim WB As Workbook

'Wype³nienie ComboBox'a cbo2ndWks nazwami arkuszy z drugiego, wybranego do porównania Workbooka.
With frmCompWks.cbo2ndWB
    'Check, czy przypadkiem nie zaznaczono pustego wiersza w ListBoxie.
    If .List(0) = vbNullString Then
       Exit Sub
    End If
      
    Set WB = Workbooks(.List(0))
    frmCompWks.cbo2ndWks.Clear
    
    i = 0
    For Each WS In Workbooks(.List(0)).Worksheets
        frmCompWks.cbo2ndWks.AddItem WS.Name, i
        i = i + 1
    Next
    frmCompWks.cbo2ndWks.ListIndex = 0
End With
End Sub

Private Sub cbo1stWks_Change()
    SprawdzCombosy
End Sub

Private Sub cbo2ndWks_Change()
    SprawdzCombosy
End Sub

Private Sub SprawdzCombosy()
'Dopisaæ sprawdzanie ListBoxa!
  Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, i As Long
  With frmCompWks
    i1 = cbo1stWks.ListIndex
    i2 = cbo2ndWks.ListIndex
    i3 = cbo1stWB.ListIndex
    For i = 0 To .cbo2ndWB.ListCount - 1
        If .cbo2ndWB.Selected(i) Then
            i4 = .cbo2ndWB.ListIndex
        End If
    Next i

    If i1 >= 0 And i2 >= 0 And i3 >= 0 And i4 >= 0 Then
      .cmdOK.Enabled = True
    Else
      .cmdOK.Enabled = False
    End If
  End With
End Sub

Private Sub cboAllTabs_Click()
    With frmCompWks
        .cbo1stWks.Enabled = Not (.cbo1stWks.Enabled)
        .cbo2ndWks.Enabled = Not (.cbo2ndWks.Enabled)
        .cboChooseRaport.Enabled = Not (.cboChooseRaport.Enabled)
        If .cboChooseRaport.Value = 1 Then
        .cboHeader.Enabled = Not (.cboHeader.Enabled)
            If .cboHeader.Value = True Then
                .cboHeader.Value = False
                .cboHeader.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub cboChooseRaport_Click()
    If cboChooseRaport.Value = 1 Then
        If cboHeader.Value = True Then
            cboHeader.Enabled = False
            'cboHeader.Value = True
        Else
            cboHeader.Enabled = True
        End If
    Else
        cboHeader.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
  With Me
    .Tag = "True"
    End
  End With
End Sub

Private Sub cmdOK_Click()
    With Me
        .Tag = "False"
        .Hide
    End With
End Sub

Private Sub UserForm_Click()

End Sub
