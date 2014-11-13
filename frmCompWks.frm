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

Private Sub cboActiveWB_Change()
  SprawdzCombosy
End Sub

Private Sub cboActiveWks_Change()
  SprawdzCombosy
End Sub

Private Sub cbo2ndWks_Change()
  SprawdzCombosy
End Sub

Private Sub SprawdzCombosy()
'Dopisaæ sprawdzanie ListBoxa!
  Dim i1 As Long, i2 As Long, i3 As Long
  With Me
    i1 = cboActiveWks.ListIndex
    i2 = cbo2ndWks.ListIndex
    i3 = cboActiveWB.ListIndex
    If i1 >= 0 And i2 >= 0 And i3 >= 0 Then
      .cmdOK.Enabled = True
    Else
      .cmdOK.Enabled = False
    End If
  End With
End Sub

Private Sub cboAllTabs_Click()
    With Me
        .cboActiveWks.Enabled = Not (.cboActiveWks.Enabled)
        .cbo2ndWks.Enabled = Not (.cbo2ndWks.Enabled)
        .cboChooseRaport.Enabled = Not (.cboChooseRaport.Enabled)
    End With
End Sub

Private Sub cmdCancel_Click()
  With Me
    .Tag = "True"
    .Hide
  End With
End Sub

Private Sub cmdOK_Click()
    With Me
        .Tag = "False"
        '.Height = 334
        .Hide
    End With
End Sub

Private Sub cbo2ndWB_Click()
Dim i As Long
Dim sPomoc As String
Dim WS As Worksheet
Dim WB As Workbook

'Wype³nienie ComboBox'a cbo2ndWks nazwami arkuszy z drugiego, wybranego do porównania Workbooka.
With Me.cbo2ndWB
    For i = 0 To .ListCount - 1
          If .Selected(i) Then
            sPomoc = .List(i)
          End If
    Next i
      
    'Check, czy przypadkiem nie zaznaczono pustego wiersza w ListBoxie.
    If sPomoc = vbNullString Then
       Exit Sub
    End If
      
    Set WB = Workbooks(sPomoc)
    frmCompWks.cbo2ndWks.Clear
    
    i = 0
    For Each WS In WB.Worksheets
        frmCompWks.cbo2ndWks.AddItem WS.Name, i
        i = i + 1
    Next
End With
End Sub
