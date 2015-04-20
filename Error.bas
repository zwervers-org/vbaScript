Attribute VB_Name = "Error"
Function LogFileName() As String

TempFolder = Environ("Temp")
ErrorFile = "ErrorLog-CertAfl" & Format(Now(), "ddmmyy") & ".txt"
LogFileName = TempFolder & "\" & ErrorFile

End Function

Private Function CDOinfo(WhatInfo)

SubName = "'CDOinfo'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim FnctionEnd

Select Case WhatInfo
    Case "sendusing"
        FnctionEnd = 2
    Case "smtpserver"
        'FnctionEnd = "mail.lieskebethke.nl"
        FnctionEnd = "smtp.gmail.com"
    Case "smtpserverport"
        'FnctionEnd = 587
        FnctionEnd = 465
    Case "smtpauthenticate"
        FnctionEnd = 1
    Case "sendusername"
        'FnctionEnd = "webbeheerder@lieskebethke.nl"
        FnctionEnd = "webbeheerder@zwervers.org"
    Case "sendpassword"
        'FnctionEnd = "9aveMxFY"
        FnctionEnd = "Zw.orgwb85"
'Gmail sender update
    Case "smtpusessl"
        FnctionEnd = True
    Case "smtpconnectiontimeout"
        FnctionEnd = 60
    Case Else
        FnctionEnd = False
        Debug.Assert FnctionEnd = False
End Select

'--------End Function

Error.DebugTekst Tekst:="Finished > " & WhatInfo & ": " & FnctionEnd, FunctionName:=SubName

OCDinfo = FnctionEnd
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

'--------Script als het niet goed gaat via de functie
    '.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    '.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webbeheerder@zwervers.org"
    '.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Zw.orgwb85"
'--------

End Function

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
    
    Error.DebugTekst Tekst:="Error values: " & vbNewLine _
                        & "->Counter: " & Counter & vbNewLine _
                        & "->SubName: " & SubName & vbNewLine _
                        & "->Erl: " & Erl & vbNewLine _
                        & "->Err.Number: " & Err.Number & vbNewLine _
                        & "->Err.Source: " & Err.Source & vbNewLine _
                        & "->Err.Description: " & Err.Description, _
                        FunctionName:="SeeText"
    
15 'back to the sheet were the error is indicated
    If ActiveSht <> "DATA" Then Sheets("DATA").Visible = xlSheetHidden
    Sheets(ActiveSht).Select
    
20  'Send an email to the opporator/bugfix-er
    Error.SendError Counter, SubName, Msg, versie
    
30  Answer = MsgBox(Msg, vbQuestion + vbOKCancel, "Error", Err.HelpFile, Err.HelpContext)
    
    If Answer = vbCancel Then End

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:="SeeText"
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next
End Function


Private Function SendError(Counter As Integer, FunctionName As String, _
                            Problem As String, versie As String)

SubName = "'SendError'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim iMsg As Object
Dim iConf As Object

40
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

Dim objBP As Object

iConf.Load -1    ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = Error.CDOinfo("sendusing")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Error.CDOinfo("smtpserver")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Error.CDOinfo("smtpserverport")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Error.CDOinfo("smtpauthenticate")
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Error.CDOinfo("sendusername")
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Error.CDOinfo("sendpassword")
    
    'Gmail sender update
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Error.CDOinfo("smtpusessl")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Error.CDOinfo("smtpconnectiontimeout")
    .Update
End With

50
   EmailTo = "anko@zwervers.org"
   'EmailFrom = "noreply@zwervers.org <noreply@zwervers.org>"
   EmailFrom = "webbeheerder@zwervers.org"
51
   EmailSubject = "Probleem in: " & ThisWorkbook.Name & " > " & FunctionName
52
   BodyText = "<font size=10px color=#FF0000>Er is een probleem gevonden in: <b>" & ThisWorkbook.Name & "</b></font></p>" _
            & "<p>Het probleem doet zich voor in de functie: <b>" & FunctionName & "</b></p>" _
            & "<p>Probleem beschrijving: </b>" & "</p><p><b>" _
            & Problem & "</b></p>"

53
   LogFile = LogFileName
54
   If Dir(LogFile) = "" Then LogFile = ""

70
With iMsg
    Set .Configuration = iConf
    .To = EmailTo
    .CC = EmailCopy
    .BCC = EmailBCC
    .From = EmailFrom
    .Subject = EmailSubject
    .HTMLBody = BodyText
    .AddAttachment LogFile
    .Send
End With

'------------
71 'Check if the problem is mentioned for the first time -> place the error in the bug tracking list
If Counter = 1 Then
    EmailTo = "x+27677550197938@mail.asana.com; anko@zwervers.org"
    'EmailFrom = "anko@zwervers.org <anko@zwervers.org>"
    EmailFrom = "webbeheerder@zwervers.org"
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
        .AddAttachment LogFile
        .Send
    End With
    
    DebugTekst "Asana email send", SubName
End If
    
'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next
End Function

Function DebugTekst(Tekst As String, Optional ByVal FunctionName As String, Optional AutoText As Boolean)

Dim s As String
Dim n As Integer

ErrorLog = Error.LogFileName

n = FreeFile()
If Dir(ErrorLog) <> "" Then
    Open ErrorLog For Append As #n
Else
    Open ErrorLog For Output As #n
End If

If IsEmpty(AutoText) Or AutoText = False Then _
    If FunctionName <> "" Then Tekst = FunctionName & ">" & Tekst

Debug.Print "--" & Format(Now(), "dd-mm-yyyy hh:mm.ss") & vbNewLine & Tekst ' write to immediate
Print #n, vbNewLine & "----" & Format(Now(), "dd-mm-yyyy hh:mm.ss") & vbNewLine & Tekst ' write to file

Close #n

End Function

Sub SendErrorLog()

SubName = "'SendErrorLog'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim iMsg As Object
Dim iConf As Object

40
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

Dim objBP As Object

iConf.Load -1    ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = Error.CDOinfo("sendusing")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Error.CDOinfo("smtpserver")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Error.CDOinfo("smtpserverport")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Error.CDOinfo("smtpauthenticate")
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Error.CDOinfo("sendusername")
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Error.CDOinfo("sendpassword")
    
    'Gmail sender update
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Error.CDOinfo("smtpusessl")
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Error.CDOinfo("smtpconnectiontimeout")

    .Update
End With

50
   EmailTo = "anko@zwervers.org"
   'EmailFrom = "noreply@zwervers.org <noreply@zwervers.org>"
   EmailFrom = "webbeheerder@zwervers.org"
51
   EmailSubject = "Error log van: " & ThisWorkbook.Name & " > " & FunctionName
52
   BodyText = "<font size=10px color=#FF0000>Logbestand van: <b>" & ThisWorkbook.Name & "</b></font></p>"
   
53

   LogFile = LogFileName()
54
   If Dir(LogFile) = "" Then LogFile = ""

70
With iMsg
    Set .Configuration = iConf
    .To = EmailTo
    .CC = EmailCopy
    .BCC = EmailBCC
    .From = EmailFrom
    .Subject = EmailSubject
    .HTMLBody = BodyText
    .AddAttachment LogFile
    .Send
End With

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Sub
