Attribute VB_Name = "Error"
Function View(ByVal What As String) As Boolean

Select Case What
'View error message
Case "Errr"
View = True
'View = False

'View screen update
Case "Updte"
'View = True
View = False

'View Alerts
Case "Alrt"
'View = True
View = False

Case ""
View = False

MsgBox "View function is empty, now set to: " & View

End Select
End Function

Function SeeText(SubName As String)
Dim Msg As String
Dim Counter As Integer
Dim versie As String

1
    Msg = "Error # " & Str(Err.Number) & Chr(13) _
            & SubName & " genarated a error. Source: " & Err.Source & Chr(13) _
            & "Error Line: " & Erl & Chr(13) _
            & Err.Description

10    'notice the error in the "Bugfix indicator"
    ActiveSht = ActiveSheet.Name
    
    If ActiveSht <> "DATA" Then Sheets("DATA").Visible = xlSheetVisible
    
    With Sheets("DATA")
        .Select
        versie = .Range("T20").Value
        Counter = .Range("T21").Value
        If .Range("T21").Value = "" Then
            Counter = 1
            .Range("T21").Value = Counter
            .Range("T22").Value = SubName
            .Range("T23").Value = Erl
            .Range("T24").Value = Err.Number
            .Range("T25").Value = Err.Source
            .Range("T26").Value = Err.Description
            Error.DebugTekst "New error in Bugfix indicator"
            
11      ElseIf .Range("T21").Value > 0 Then
            If SubName = "'" & .Range("T22").Value And _
                Erl = .Range("T23").Value And _
                Err.Number = .Range("T24").Value And _
                Err.Source = .Range("T25").Value Then
                    .Range("T21").Value = Counter + 1
12          Else
                Counter = 1
                .Range("T21").Value = Counter
                .Range("T22").Value = SubName
                .Range("T23").Value = Erl
                .Range("T24").Value = Err.Number
                .Range("T25").Value = Err.Source
                .Range("T26").Value = Err.Description
                Error.DebugTekst "Delete previous and add new error in Bugfix indicator"
            End If
        End If
    Counter = .Range("T21").Value
    End With
    
15 'back to the sheet were the error is indicated
    If ActiveSht <> "DATA" Then Sheets("DATA").Visible = xlSheetHidden
    Sheets(ActiveSht).Select
    
20  'Send an email to the opporator/bugfix-er
    Error.SendError Counter, SubName, Msg, versie
    
30  Answer = MsgBox(Msg, vbQuestion + vbOKCancel, "Error", Err.HelpFile, Err.HelpContext)
    
    If Answer = vbCancel Then End
    
End Function


Private Function SendError(Counter As Integer, FunctionName As String, _
                            Problem As String, versie As String)

SubName = "'SendError'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Dim iMsg As Object
Dim iConf As Object

40
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

Dim objBP As Object

iConf.Load -1    ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
                   = "mail.lieskebethke.nl"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webbeheerder@lieskebethke.nl"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "9aveMxFY"
    .Update
End With

50
   EmailTo = "anko@zwervers.org"
   EmailFrom = "noreply@zwervers.org <noreply@zwervers.org>"
51
   EmailSubject = "Probleem in: " & ThisWorkbook.Name & " > " & FunctionName
52
   BodyText = "<font size=10px color=#FF0000>Er is een probleem gevonden in: <b>" & ThisWorkbook.Name & "</b></font></p>" _
            & "<p>Het probleem doet zich voor in de functie: <b>" & FunctionName & "</b></p>" _
            & "<p>Probleem beschrijving: </b>" & "</p><p><b>" _
            & Problem & "</b></p>"

70
With iMsg
    Set .Configuration = iConf
    .To = EmailTo
    .CC = EmailCopy
    .BCC = EmailBCC
    .From = EmailFrom
    .Subject = EmailSubject
    .HTMLBody = BodyText
    .Send
End With

'------------
71 'Check if the problem is mentioned for the first time -> place the error in the bug tracking list
If Counter = 1 Then
    EmailTo = "x+27677550197938@mail.asana.com"
    EmailFrom = "anko@zwervers.org <anko@zwervers.org>"
    EmailSubject = "v" & versie & " > " & FunctionName
    
    BodyText = "<p>Probleem beschrijving: </b>" & "</p><p><b>" _
                    & Problem & "</b></p>"

    With iMsg
        Set .Configuration = iConf
        .To = EmailTo
        .CC = ""
        .BCC = ""
        .From = EmailFrom
        .Subject = EmailSubject
        .HTMLBody = BodyText
        .Send
    End With
    
    DebugTekst "Asana email send", SubName
End If
'-----------

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function DebugTekst(Tekst As String, Optional ByVal FunctionName As String)

If Not IsEmpty(FunctionName) Then Tekst = Left$(FunctionName & ":" & Space(Padding), Padding) & Tekst

Debug.Print Tekst

End Function

