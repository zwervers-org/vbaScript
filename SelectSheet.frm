VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectSheet 
   Caption         =   "Select Sheet"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4290
   OleObjectBlob   =   "SelectSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GoAction_Click()

SubName = "'GoAction_Click'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

1
Admin.ShowAllSaveSheet

10
With SheetsSelected
    For i = 0 To .ListCount - 1
        If .Selected(i) = True Then
20
            Select Case SelectSheet.ActivateAction.Caption
                Case "Delete sheets"
30
                    Worksheets(.List(i)).Delete
                Case Else
90
                    Error.DebugTekst ("No action selected for sheet: " & .List(i))
                    GoTo ExitSub
            End Select
100
            Error.DebugTekst (SelectSheet.ActivateAction.Caption & ": " & .List(i))
        End If
    Next i
    BackgroundFunction.AutoCloseMessage Tekst:=SelectSheet.ActivateAction.Caption
End With

ExitSub:
110
SelectSheet.Hide
Admin.HideAllSaveSheets

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub


Private Sub UserForm_Initialize()

With SheetsSelected
    For i = 1 To Sheets.count
        If BackgroundFunction.InArray("Sheets", Sheets(i).Name) = False Then _
            .AddItem Sheets(i).Name
    Next i
End With

End Sub
