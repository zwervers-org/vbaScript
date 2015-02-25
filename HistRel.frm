VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HistRel 
   Caption         =   "Zoek geschiedenis van relatie"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   OleObjectBlob   =   "HistRel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HistRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GoSearch_Click()

HistRel.Hide

If HistRel.NeddoxCode.Value = "" And HistRel.SearchCertificate.Value = "" Then
    MsgBox "Put a search value in one of the boxes"
    HistRel.Show
Else
    
    If HistRel.NeddoxCode.Visible = True Then
        RelHistory.SearchHistory code:=HistRel.NeddoxCode.Value, SearchType:="Neddox"
    
    ElseIf HistRel.SearchCertificate.Visible = True Then
        RelHistory.SearchHistory code:=HistRel.SearchCertificate.Value, SearchType:="Certificate"
        
    Else
        MsgBox "No value to search"
        HistRel.Show
    End If
    
    ClearCertList
    
    Admin.ShowOneSheet ("Certificaten")
    
End If

End Sub


Private Sub SearchCertificate_Change()

If HistRel.SearchCertificate.Value <> "" Then
    HistRel.NeddoxCode.Visible = False
    HistRel.Neddox.Visible = False
Else
    HistRel.NeddoxCode.Visible = True
    HistRel.Neddox.Visible = True
End If

End Sub

Private Sub NeddoxCode_Change()

If HistRel.NeddoxCode.Value <> "" Then
    HistRel.SearchCertificate.Visible = False
    HistRel.Certificate.Visible = False
Else
    HistRel.SearchCertificate.Visible = True
    HistRel.Certificate.Visible = True
End If

End Sub

Private Sub UserForm_Activate()

If Selection.Column = 3 And Selection.count = 1 And Sheets("Certificaten") Is ActiveSheet Then
    
    HistRel.Hide
    
    RelHistory.SearchHistory code:=ActiveCell.Value, SearchType:="Neddox"

Else
    'Zet alles in de selectielijst voor certificaten
    Admin.ShowOneSheet ("DATA")
    
    Dim Counter As Integer
    
    Counter = 2 'startrow of certificatetypes summary
    
    With Sheets("DATA")
        ClearCertList
        
        EindeCert = .Range("Z2").End(xlDown).Row 'einde voor aantal certificaat-typen resetten
        EindeBio = .Range("AA2").End(xlToRight).Column 'einde voor aantal bio-typen resetten

        For Cert = 2 To EindeCert
            .Range("Y" & Counter).Value = .Range("Z" & Cert).Value
            Counter = Counter + 1
        Next Cert
        
        For Bio = 27 To EindeBio
            .Range("Y" & Counter).Value = .Cells(2, Bio).Value
            Counter = Counter + 1
        Next Bio
    End With
    
    LastCertRow = Sheets("DATA").Range("Y2").End(xlDown).Row

    HistRel.SearchCertificate.RowSource = "DATA!Y2:Y" & LastCertRow
    
    Admin.ShowOneSheet ("Certificaten")
End If

End Sub


Private Sub UserForm_Terminate()

ClearCertList

Admin.ShowOneSheet ("Certificaten")

Range("A2").Select

End Sub

Private Sub ClearCertList()

'Admin.ShowOneSheet ("DATA")
With Sheets("DATA")
    .Range("Y1:Y" & Sheets("DATA").Range("Y2").End(xlDown).Row).ClearContents
End With

End Sub
