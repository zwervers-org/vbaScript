VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HistorySearch 
   Caption         =   "Results history search"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13770
   OleObjectBlob   =   "HistorySearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HistorySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CloseWithoutButton_Click()

Unload Me

RelHistory.CloseOverview ("No")

End Sub

Private Sub CloseWithSaveButton_Click()

Unload Me

RelHistory.CloseOverview ("Yes")

End Sub

Private Sub CloseButton_Click()

Unload Me

RelHistory.CloseOverview ("Do you whant to keep this information?")

End Sub

Private Sub ListBox1_Click()
'De waarde van de regel weergeven in een berichtbox

Dim Msg As String
Dim i As Integer

Msg = "Informatie:" & vbNewLine

For i = 0 To HistorySearch.ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        Sht = "Sheet: " & ListBox1.Column(0, i) & vbNewLine
        code = "Code: " & ListBox1.Column(1, i) & vbNewLine
        naam = "Naam: " & ListBox1.Column(2, i) & vbNewLine
        div = "Divisie: " & ListBox1.Column(3, i) & vbNewLine
        actn = "Actie: " & ListBox1.Column(4, i) & vbNewLine
        dte = "Geldig tot: " & ListBox1.Column(5, i) & vbNewLine
        Cert = "Certificaat: " & ListBox1.Column(6, i) & vbNewLine
        commnt = "Opmerking: " & ListBox1.Column(7, i) & vbNewLine
        
        Msg = Msg & Sht & div & naam & dte & Cert & actn & commnt
    End If
Next i
MsgBox Msg

End Sub

Private Sub UserForm_Activate()

    DoEvents
    
    Call RelHistory.PublishHistorySearch

End Sub

