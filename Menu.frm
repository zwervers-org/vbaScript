VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Menu"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CleanSht_Click()

Menu.Hide

CleanCert

End Sub


Private Sub CommandButton1_Click()

Menu.Hide

Call SavePDF

BackgroundFunction.AutoCloseMessage Tekst:="Saved! to j:\Certificaten\" & Range("A1") & ".pdf"

End Sub

Private Sub AutoFillActions_Click()

Menu.Hide

Call FillActions

End Sub

Sub SearchHist_Click()

Menu.Hide

HistRel.Show

End Sub

Private Sub UserForm_Activate()

DateValidCert.Caption = Sheets("Certificaten").Range("A1").Value

End Sub

Private Sub UserForm_Initialize()

Dim i As Integer

With GotoSht
For i = 1 To Sheets.count

.AddItem Sheets(i).Name

Next i
End With

End Sub

Private Sub SaveData_Click()

Menu.Hide

SaveOldData

BackgroundFunction.AutoCloseMessage Tekst:="Saved!"

End Sub


Private Sub Email1_Click()

Menu.Hide

SorterenEmail ("Aanvragen")

End Sub

Private Sub Email2_Click()

Menu.Hide

SorterenEmail ("Email")

End Sub

Private Sub FilterInkoper_Click()

Menu.Hide

SortInkoper.InkoperSorteren

BackgroundFunction.AutoCloseMessage Taak:="SortInkoper"

End Sub

Private Sub LoadLdData_Click()

Menu.Hide

LoadData.LoadOldData ("")

End Sub

Private Sub LoadNwData_Click()

Menu.Hide

OpenFile

End Sub

Private Sub GotoSht_Click()

Menu.Hide

ShowOneSheet (GotoSht.Value)

End Sub

Private Sub GotoSht_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Menu.Hide

ShowOneSheet (GotoSht.Value)

End Sub
