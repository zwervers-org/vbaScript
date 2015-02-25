Attribute VB_Name = "ClearData"
Sub ClearSheet(Sht As String)

SubName = "'ClearSheet'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

1 If Sht = "Certificaten" Then

'Clear data in Certificaten

10 Sheets("Certificaten").Select

11 CertBewerkbaar

'Clear existing data

20    Eind = Range("C1000").End(xlUp).Row

21    If Eind > 1 Then

22    Set rng = Range("A2", "G" & Eind)
    
23    rng.ClearContents
    
'Second clear existing data
30    Sheets("Certificaten").Select
    
31    Set rng = Range("L2", "L" & Eind)
    
32    rng.ClearContents
    
'Clear date
33    Range("A1").ClearContents
    
'Set back to workmode
40    CertNietBewerkbaar

29    End If
    
100 Else

110 If Not InArray("Sheets", Sht) Then

'Start with clean sheet to save the data in
111    Sheets(Sht).Visible = xlSheetVisible
112    Sheets(Sht).Select
    
113    EindeData = Range("C1000").End(xlUp).Row
    
120    If EindeData > 1 Then

121    Range("A2", "L" & EindeData).ClearContents
129    End If
119 End If
999 End If
Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Sub CleanCert()

SubName = "'CleanCert'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Dim rw As Range
Dim rng As Range
Dim CleanSkp As Integer
Dim CleanAmnt As Integer

CleanSkp = 0
CleanAmnt = 0

1   CertBewerkbaar

8   EindeCert = Range("E1000").End(xlUp).Row

10        Set rng = Range("A2", "L" & EindeCert)
11        For Each rw In rng.Rows
20            If Range("G" & rw.Row).Value = "" Then
                If Range("A" & rw.Row).Value = "" Then
25                    Range("A" & rw.Row, "G" & rw.Row).ClearContents
26                    Range("L" & rw.Row, "L" & rw.Row).ClearContents
27                      CleanAmnt = CleanAmnt + 1

21                Else
                application.ScreenUpdating = True
                ActiveWindow.ScrollRow = rw.Row
                
                MValue = MsgBox("Needs row: " & rw.Row & " to be cleaned?", vbYesNo, "Is this correct?")
                application.ScreenUpdating = False
                
30                    If MValue = vbNo Then
31                        CleanSkp = CleanSkp + 1
                            Range("G" & rw.Row).Value = " "
35                    Else
36                        Range("A" & rw.Row, "G" & rw.Row).ClearContents
37                        Range("L" & rw.Row, "L" & rw.Row).ClearContents
                        
38                        CleanAmnt = CleanAmnt + 1
39                    End If
                End If
29            End If
19        Next rw

101        For Each rw In rng.Rows
120            If Range("G" & rw.Row).Value = " " Then
121                 Range("C" & rw.Row).Value = ""
125            End If
109        Next rw

9   EindeCert = Range("E1000").End(xlUp).Row

'sorteren op naam
40  Sheets("Certificaten").Sort.SortFields.Clear
41  Sheets("Certificaten").Sort.SortFields.Add Key:=Range( _
        "C2", "C" & EindeCert + 100), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

50  With ActiveWorkbook.Worksheets("Certificaten").Sort
        .SetRange Range("A1", "L" & EindeCert)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
59  End With

'Certificaat sheet beveiligen
60  CertNietBewerkbaar

61  ActiveWindow.ScrollRow = 2
62  Range("A2").Select

'bericht weergeven en automatisch wegklikken
73  BackgroundFunction.AutoCloseMessage Tekst:="There are " & CleanAmnt & " cleaned, and " & CleanSkp & " skiped."

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub
