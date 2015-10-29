Sub Mail_Range()

'Mean sub to start the email sending
SubName = "Mail_Range"
On Error GoTo ErrorText:

Application.ScreenUpdating = False
Application.DisplayAlerts = False

DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

    Dim OutApp As Object
    Dim OutMail As Object
    'Dim OutAtt As Object
    Dim FindString As String
    Dim rng As Range
    Dim rw As Integer
    Dim clmn As Integer
    Dim Clm As Integer
    Dim MailSubject As String
    Dim MailBody As String
    Dim AddPDF As String
    Dim PriceListRange As Range
30
'---------You can change the information below

'Texts need to start AND end with: "
'A new line can be started with the 'VbNewLine', If you need to connect the text and new line use a '&' character
'Example: "This is the text" & VbNewLine & "This is text on a new line"

'Subject of email question
    MailSubjectTxt = "Geef het ontwerp van de email." 'Text
    MailSubjectTitle = "Email onderwerp" 'Title

'Question if the selected range must be added in the body of the email
    MailBodyTxt = "Prijslijst opnemen in de email zelf?" 'Text
    MailBodyTitle = "Prijslijst in email" 'Title

'Question if the selected range must be added as attachement to the email
    AddPdfTxt = "Prijslijst als PDF toevoegen aan de email?" 'Text
    AddPdfTitle = "Prijslijst als PDF toevoegen aan de email?" 'Title

'Save information for later question
    SaveInfoTxt = "Informatie opslaan voor later gebruik?" & vbNewLine & vbNewLine & "PS. Dit gebeurt alleen voor het huidige tabblad"
    SaveInfoTitle = "Informatie opslaan?"

'Question to select the range you want to use in the email
    PriceListRangeTxt = "Selecteer het gebied van de prijslijst" 'Text
    PriceListRangeTitle = "Selecteer prijslijst" 'Title

'Show saved information alert box
    SavedInfoTxt = "Informatie die is opgeslagen gebruiken?" & vbNewLine & "Opgeslagen informatie:"
    SavedInfoTitle = "Opgeslagen informatie"

'--------The script starts here, DO NOT CHANGE HERE IF YOU NOT ARE AN EXPERT

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

40
'When information is saved, skip questions and show saved information
If IsEmpty(ActiveSheet.Range("ZY1")) And IsEmpty(ActiveSheet.Range("ZZ3")) Then
Questions:
    MailSubject = InputBox(Prompt:=MailSubjectTxt, Title:=MailSubjectTitle)
    If MailSubject = "" Then End
    
    MailBody = MsgBox(Prompt:=MailBodyTxt, Title:=MailBodyTitle, Buttons:=vbYesNo)
    
    AddPDF = MsgBox(Prompt:=AddPdfTxt, Title:=AddPdfTitle, Buttons:=vbYesNo)
    
    SaveInfo = MsgBox(Prompt:=SaveInfoTxt, Title:=SaveInfoTitle, Buttons:=vbYesNo)


    Application.ScreenUpdating = True
    Set PriceListRange = Application.InputBox(Prompt:=PriceListRangeTxt, Title:=PriceListRangeTitle, Type:=8, Left:=50, Top:=50)
    If PriceListRange Is Nothing Then End
    Application.ScreenUpdating = False
Else
45
'When there is information saved
    With ActiveSheet
        MailSubject = .Range("ZZ2").Value
        MailBody = .Range("ZZ3").Value
        AddPDF = .Range("ZZ4").Value
        Set PriceListRange = Range(.Range("ZZ5").Value)
        
    End With
    
    If MailBody = 6 Then
        MailBody1 = "Yes"
    Else
        MailBody1 = "No"
    End If
    
    If AddPDF = 6 Then
        AddPDF1 = "Yes"
    Else
        AddPFD1 = "No"
    End If

'Show saved information in message box
46
    SavedInfo = MsgBox(Prompt:=SavedInfoTxt & vbNewLine _
                & "    Subject email: " & MailSubject & vbNewLine _
                & "    Selection in email:   " & MailBody1 & vbNewLine _
                & "    Selected area as attachment: " & AddPDF1 & vbNewLine _
                & "    Selected area:     " & PriceListRange.Address, _
                Title:=SavedInfoTitle, Buttons:=vbYesNo)
                    
    If SavedInfo = vbYes Then
        GoTo SkipQuestions
    Else
        ActiveSheet.Range("ZY1:ZZ10").Clear
        GoTo Questions
    End If
End If

50
'Save information to range
If SaveInfo = vbYes Then
    With ActiveSheet
        .Range("ZY1").Value = "Saved mail information:"
        .Range("ZY2").Value = "MailSubject:"
        .Range("ZY3").Value = "Add pricelist in body?"
        .Range("ZY4").Value = "Add pricelist as PDF?"
        .Range("ZY5").Value = "Pricelist range:"
       'DATA
        .Range("ZZ2").Value2 = MailSubject
        .Range("ZZ3").Value2 = MailBody
        .Range("ZZ4").Value2 = AddPDF
        .Range("ZZ5").Value2 = PriceListRange.Address
    End With
End If

SkipQuestions:
55
'Save range in HTML format
If MailBody = 6 Then 'vbYes
    MailBody = RangetoHTML(PriceListRange)
Else
    MailBody = ""
End If

56
'Save range in PDF
If AddPDF = 6 Then 'vbYes
    AddPDF = SavePDF(PriceListRange)
Else
    AddPDF = ""
End If
            
60
With OutMail
    .To = ""
    .CC = ""
    .BCC = ""
                    
    .Subject = MailSubject
    .HTMLBody = MailBody
    .Attachments.Add AddPDF
    .Display   'or use .Send for direct send
End With

'Save AddPDF file name
With ActiveSheet
    .Range("ZY6").Value2 = "File name and folder PDF"
    .Range("ZZ6").Value2 = AddPDF
End With

65 Set OutMail = Nothing
66 Set OutApp = Nothing
67 If AddPDF <> "" Then Kill AddPDF

'--------End Function
DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Sub

Private Function SavePDF(TxtRange As Range) As String

'Save selected range into a PDF file

SubName = "SavePDF"
On Error GoTo ErrorText:

Application.ScreenUpdating = False
Application.DisplayAlerts = False
DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function
Dim FileName As String

10
TempFolder = Environ("Temp")
FileName = TempFolder & "\" & ActiveSheet.Name & Format(Now(), "dd-mm-yyyy") & ".pdf"
20
TxtRange. _
        ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        FileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
30
SavePDF = FileName

'--------End Function
DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function

Private Function SeeText(SubName As String)
Dim Msg As String

'This function gives a error warning in an alert with option to go further or hold

10 'This must be add first in the function otherwise the error information is cleared
    Msg = "Error # " & Str(Err.Number) & Chr(13) _
            & SubName & " genarated a error. Source: " & Err.Source & Chr(13) _
            & "Error Line: " & Erl & Chr(13) _
            & Err.Description & Chr(13) _
            & "Please send this error incl. log file to: webbeheerder@zwerver.org & Chr(13)" _
            & "The log file can be found here: " & LogFileName
20
    DebugTekst Tekst:="Error values: " & vbNewLine _
                        & "->Counter: " & Counter & vbNewLine _
                        & "->SubName: " & SubName & vbNewLine _
                        & "->Erl: " & Erl & vbNewLine _
                        & "->Err.Number: " & Err.Number & vbNewLine _
                        & "->Err.Source: " & Err.Source & vbNewLine _
                        & "->Err.Description: " & Err.Description, _
                        FunctionName:="SeeText"
    
30  Answer = MsgBox(Msg, vbQuestion + vbOKCancel, "Error", Err.HelpFile, Err.HelpContext)
    
    If Answer = vbCancel Then End

'--------End Function
DebugTekst Tekst:="Finish", FunctionName:="SeeText"
Exit Function

ErrorText:
If Err.Number <> 0 Then
    MsgBox "Fout in ErrorHandler"
    End
End If

Resume Next
End Function

Private Function DebugTekst(Tekst As String, Optional ByVal FunctionName As String, Optional AutoText As Boolean)

'This function makes an errorlog in your temp directory, and shows the task in the statusbar of Excel

Dim s As String
Dim n As Integer
On Error Resume Next

ErrorLog = LogFileName

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

If IsEmpty(AutoText) Or AutoText = False Then
    If FunctionName <> "" Then Application.StatusBar = FunctionName & "> " & Tekst
Else
    If FunctionName <> "" Then Application.StatusBar = FunctionName & "> " & Tekst
End If

Close #n

End Function

Function LogFileName() As String

TempFolder = Environ("Temp")
ErrorFile = ThisWorkbook.Name & Format(Now(), "ddmmyy") & ".err"
LogFileName = TempFolder & "\" & ErrorFile

End Function

Private Function RangetoHTML(rng As Range)

'This puts the selected area in a HTML format to add as body in the email (without pictures)

SubName = "RangeToHtml"
On Error GoTo ErrorText:

Application.ScreenUpdating = False
Application.DisplayAlerts = False
DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook
   
1    TempFile = Environ$("temp") & "\TempPriceList.htm"

    'Copy the range and create a new workbook to past the data in
2    rng.Copy
3    Set TempWB = Workbooks.Add(1)
10    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).PasteSpecial
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo ErrorText:
19    End With

    'Publish the sheet to a htm file
20    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
29    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
30    RangetoHTML = ts.readall
31    ts.Close
32    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
33    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
34    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing

'--------End Function
DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)

'To delete tempory files add closing the file

ErrorLog = LogFileName
AddPDF = Range("ZZ6").Value

If ErrorLog <> "" Then If Dir(ErrorLog) <> "" Then Kill ErrorLog

If AddPDF <> "" Then If Dir(AddPDF) <> "" Then Kill AddPDF

End Sub


