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

Private Sub MultiPage1_Change()

Dim RngToetsCombi As Range
Dim RngCelCombi As Range

1
'If MultiPage1.Pages("Help").Visible = True Then
If MultiPage1.SelectedItem.Name = "Help" Then

    'clean lists and declare variables
    HotKeys1.Clear
    CelKeys1.Clear
    regel = 0
    ActSht = ActiveSheet.Name
    
10
    If ActiveSheet.Name <> "DATA" Then Admin.ShowOneSheet ("DATA")
20
    With Sheets("DATA")
        Set RngToetsCombi = .Range("S37:X" & .Range("S50").End(xlUp).Row)
        Set RngCelCombi = .Range("S51:X" & .Range("S64").End(xlUp).Row)
30
        'fill list Key combinations
        For Each rw In RngToetsCombi
            With HotKeys1
                .ColumnHeads = True
                .ColumnCount = 2
                .RowSource = RngToetsCombi.Address
            End With
        Next rw
40
        'fill list cel actions
        For Each rw In RngCelCombi
            With CelKeys1
                .ColumnHeads = True
                .ColumnCount = 2
                .RowSource = RngCelCombi.Address
            End With
        Next rw
    End With
50
    If ActSht <> "DATA" Then Admin.ShowOneSheet (ActSht)

End If

End Sub

Sub SearchHist_Click()

Menu.Hide

HistRel.Show

End Sub

Private Sub SynergyDocNr1_Onclick()

Dim IE As Object

' Create InternetExplorer Object
Set IE = CreateObject("InternetExplorer.Application")

' You can uncoment Next line To see form results
IE.Visible = False

' Send the form data To URL As POST binary request
IE.Navigate "http://synergy/docs/DocView.aspx?DocumentID={a1dfb942-2a25-463a-bf7f-5599148ee728}"

application.StatusBar = "Please wait... Explorer is opening"

' Wait while IE loading...
Do While IE.Busy
    application.Wait DateAdd("s", 1, Now)
Loop

application.StatusBar = "Explorer is opened the document"
 
' Show IE
IE.Visible = True

Set IE = Nothing

application.StatusBar = ""

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
