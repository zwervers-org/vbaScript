Attribute VB_Name = "BackgroundFunction"

Private Function SetPassword()

SetPassword = "DocCertOverview123"

End Function

Function MenuShow()

SubName = "'MenuShow'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

1 If ActiveSheet.Name <> "Certificaten" Then

2 Menu.MultiPage1("OptionsTab").Visible = False

10 Else
11 Menu.MultiPage1("OptionsTab").Visible = True
12 Menu.MultiPage1.Value = 0
End If

20 Menu.Show

Exit Function


ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function InArray(WitchArray, strValue)

SubName = "'InArray'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Dim j

If WitchArray = "" Then GoTo EndFunction
If strValue = "" Then GoTo EndFunction

10
Select Case WitchArray
Case "Sheets"
    ArrayVals = Array("DATA", "Certificaten", "Aanvragen", "Email", "SortInk", "EmailSjabloon", "NotAvailable")
Case "NotAv"
    ArrayVals = Array("1", "2", "3", "8", "9", "10", "11", "12")
Case "VBAExport"
    ArrayVals = Array(".frm", ".bas", ".txt")
Case Else
    Error.DebugTekst "No Array selected:" & WitchArray & vbNewLine & "String: " & strValue, SubName
    GoTo EndFunction
End Select

20  For j = 0 To UBound(ArrayVals)
21    If ArrayVals(j) = CStr(strValue) Then
22      InArray = True
      Exit Function
    End If
  Next
  
EndFunction:
25  InArray = False
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function CertBewerkbaar()

SubName = "'CertBewerkbaar'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", FunctionName:=SubName

1 Admin.HideAllSaveSheets
2 Sheets("Certificaten").Select
3 ActiveSheet.Unprotect Password:=SetPassword
4 Columns("D:G").EntireColumn.Hidden = False

16 If ActiveSheet.AutoFilterMode = True Then

15 Range("A:Z").AutoFilter

Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

End If

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function CertNietBewerkbaar()

SubName = "'CertNietBewerkbaar'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Dim CriteriaValue As String
Dim CriteriaSplit As String

Error.DebugTekst Tekst:="Start", FunctionName:=SubName

CriteriaSplit = "|"

1 Admin.HideAllSheets
2 Sheets("Certificaten").Select

10 EndCert = Range("C1000").End(xlUp).Row

'Rijen netjes maken
For rij = 2 To EndCert
    Rows(rij).EntireRow.AutoFit
Next rij
11 Range("A2", "L" & EndCert).Locked = False 'beveiliging uitzetten
   Range("C2", "F" & EndCert).Locked = True 'beveiliging aanzetten
13 Columns("E:F").EntireColumn.Hidden = True 'kolom verbergen
   Range("H2", "K" & EndCert).Locked = True 'beveiliging aanzetten
14 Range("M2", "ZZ" & EndCert).Locked = True 'beveiliging aanzetten

15 Range("A:Z").AutoFilter

20 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=False, AllowSorting:=True, AllowFiltering:=True, Password:=SetPassword
21 ActiveSheet.EnableSelection = xlNoRestrictions

Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Function ColLett(Col As Integer) As String
     
    If Col > 26 Then
        ColLett = ColLett((Col - (Col Mod 26)) / 26) + Chr(Col Mod 26 + 64)
    Else
        ColLett = Chr(Col + 64)
    End If
     
End Function

Function PrintArray(StrArray As Variant)
  
  For k = 0 To UBound(StrArray)
    txt = txt & k & ": " & StrArray(k) & vbCrLf
  Next k
  
  MsgBox txt

End Function

Public Function ShowFilter(rng As Range) As String

Dim oFilter As Filter
Dim sCriteria1 As String
Dim sCriteria2 As String
Dim aCriteria As String
Dim sOperator As String
Dim nOp As Long
Dim nOff As Long
Dim rngFilter As Range
Dim sh As Worksheet
Dim ColNr As Integer
Dim CriteriaSplit As String
CriteriaSplit = "|"

    Set sh = rng.Parent
    If sh.FilterMode = False Then
        Exit Function
    End If
    Set rngFilter = sh.AutoFilter.Range

    If Intersect(rng.EntireColumn, rngFilter) Is Nothing Then
        ShowFilter = CVErr(xlErrRef)
    Else
        nOff = rng.Column - rngFilter.Columns(1).Column + 1
        ColNr = nOff
        If Not sh.AutoFilter.Filters(nOff).On Then
            ShowFilter = ""
        Else
            Set oFilter = sh.AutoFilter.Filters(nOff)
            nOp = oFilter.Operator
            If nOp = xlFilterValues Then

                ShowFilter = oFilter.Criteria1
            
            Else
                On Error Resume Next
                sCriteria1 = oFilter.Criteria1
                sCriteria2 = oFilter.Criteria2
                
                'sOperator = ""
                'If nOp = xlAnd Then
                    sOperator = CriteriaSplit
                'ElseIf nOp = xlOr Then
                '    sOperator = CriteriaSplit
                'End If
            
                ShowFilter = sCriteria1 & sOperator & sCriteria2
            End If
        End If
    End If
End Function

Function AutoCloseMessage(Optional Titel As String, Optional Taak As String, Optional Interval As Integer, Optional Tekst As String, _
                            Optional VoetTekst As String)

SubName = "'AutoCloseMessage'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Dim InfoBox As Object

1 'Body tekst
If Tekst = "" Then
    Select Case Taak
        Case "SortInkoper"
            Tekst = "Sorted per buyer"
        Case "SortEmail"
            Tekst = "Prepared addresses for emailing"
        Case Else
            Tekst = "There is some task completed"
    End Select
End If

4 'plaats de tekst in de debug log
Error.DebugTekst Tekst:="Tekst: " & Tekst & vbNewLine _
                        & "Titel: " & Titel & vbNewLine _
                        & "VoetTekst: ", FunctionName:=SubName, _
                        AutoText:=True

5 'titel
If Titel = "" Then _
    Titel = "Task complete"

6 'Close time in seconds
If Interval = 0 Then _
    Interval = 2

7 'Voettekst
If VoetTekst = "" Then _
    VoetTekst = "PRESS: 'ctrl+m' FOR THE MENU" & vbNewLine & vbNewLine _
                & "(Auto close: " & Interval & "sec)."

10
Set InfoBox = CreateObject("WScript.Shell")

application.StatusBar = Titel & ": " & Tekst
20
Select Case InfoBox.Popup(Tekst & vbNewLine & vbNewLine & VoetTekst, Interval, Titel, 0)
    Case 1, -1
        Exit Function
End Select

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Function
