Attribute VB_Name = "SndEmail"
Sub FillSjabloon(Sht As String)

Dim code As String
Dim rng As Range
Dim Eind As Integer
Dim EindData As Integer
Dim rw As Range
Dim FindString As String

CertBewerkbaar

SubName = "'FillSjabloon'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

Sheets(Sht).Select

2 Eind = Range("A1000").End(xlUp).Row

'bij één contact adres alleen kopiëren
10 If Eind = 2 Then
11  Range("A" & Eind).Copy
12  Range("M2").PasteSpecial (xlPasteValues)

15 Else
'zorgen voor unieke waarden (code)
16    Range("A2", "A" & Eind).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "A2", "A" & Eind), CopyToRange:=Range("M2"), Unique:=True
        
20    If Range("M2").Value = Range("M3").Value Then
21    Range("M2").Delete Shift:=xlUp
29    End If
19 End If

3 EindData = Range("M1000").End(xlUp).Row

4 Set rng = Range("M2", "M" & EindData)

'mark send date and time
5 Range("M1").Value = Format(Now, "dd-mm-yyyy hh:mm")

30 For Each rw In rng
    
31    Sheets(Sht).Select
    
32    Cells(rw.Row, 13).Copy
33    Sheets("EmailSjabloon").Range("B6").PasteSpecial (xlPasteValues)
    
    'zien hoeveel certificaten er per code zijn
34    FindString = Cells(rw.Row, 13)
    
    
40    With Sheets(Sht).Columns(1)
41    Dim rng1 As Range
42    Dim rng2 As Range
    
43    Set rng1 = Nothing
44    Set rng2 = Nothing
45                Set rng1 = .Find(What:=FindString, _
                                After:=.Cells(1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
46                Set rng2 = .Find(What:=FindString, _
                                After:=.Cells(Eind + 1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious, _
                                MatchCase:=False)
50                If rng1.Row = rng2.Row Then
                'Er is 1 certificaat die verloopt
51                  copyrng = Range("C" & rng1.Row, "E" & rng1.Row).Address
52                  Range(copyrng).Copy
53                  Sheets("EmailSjabloon").Range("U2").PasteSpecial (xlPasteValues)
                
54                ElseIf rng1.Row < rng2.Row Then
                'Er zijn meerdere certificaten die verlopen
55                  Range("C" & rng1.Row, "E" & rng2.Row).Copy
56                  Sheets("EmailSjabloon").Range("U2").PasteSpecial (xlPasteValues)
                
57                Else
58                  BackgroundFunction.AutoCloseMessage Tekst:="There is a problem with getting the amount of expiring certificates"
59                End If
                
                'Copy other information
                    'naam
60                Range("F" & rng1.Row, "G" & rng1.Row).Copy
61                Sheets("EmailSjabloon").Range("B2").PasteSpecial (xlPasteValues)
                    'email
62                Range("H" & rng1.Row).Copy
63                Sheets("EmailSjabloon").Range("B3").PasteSpecial (xlPasteValues)
                    'CC
64                Range("J" & rng1.Row).Copy
65                Sheets("EmailSjabloon").Range("B4").PasteSpecial (xlPasteValues)
                    'taal
66                Range("I" & rng1.Row).Copy
67                Sheets("EmailSjabloon").Range("B5").PasteSpecial (xlPasteValues)
                    'datum
68                'Range("M1").Copy
                  'Sheets("EmailSjabloon").Range("L2").PasteSpecial (xlPasteValues)
                  GestuurdOpStart = InStr(Range("K" & rng1.Row).Value, "Gestuurd op: ")
                  
                  If GestuurdOpStart > 0 Then
                    SendTime = Mid(Range("K" & rng1.Row).Value, GestuurdOpStart + 13, 16)
                    If Not IsEmpty(SendTime) Then _
                        Sheets("EmailSjabloon").Range("L2").Value = SendTime
                  ElseIf Sht = "Email" Then
                    Sheets("EmailSjabloon").Range("L2").Value = Format(Now, "dd-mm-yyyy hh:mm")
                  End If
49    End With
    Mail_Range (Sht)

70    If Sheets(Sht).Range("L1").Value = False Then
71    Sheets(Sht).Range("L" & rw.Row).Value = 0
    'MsgBox "Email is not send"
    GoTo Volgende:
    
75    Else
    'Set action on email send
  
80    With Sheets("Certificaten")
        Sheets("Certificaten").Select
81        Dim rw1 As Range
82        Dim EindCert As Integer
83        Dim ARng As Range
        
        EindCert = Range("C1000").End(xlUp).Row
        
84        Set ARng = Range("A2", "C" & EindCert)
90        For Each rw1 In ARng
91            If Range("C" & rw1.Row).Value = FindString Then
92                Select Case Sht
                Case "Aanvragen"
100                    If Range("A" & rw1.Row).Value = 1 Then
101                        Range("A" & rw1.Row).Value = 2
                           If IsEmpty(Range("L" & rw1.Row).Value) Then
                                Range("L" & rw1.Row).Value = "Gestuurd op: " _
                                & Format(Now, "dd-mm-yyyy hh:mm")
                           Else
                                Range("L" & rw1.Row).Value = "Gestuurd op: " _
                                & Format(Now, "dd-mm-yyyy hh:mm") & " | " & Range("L" & rw1.Row).Value
                           End If
109                    End If
                
                Case "Email"
110                    If Range("A" & rw1.Row).Value = 2 Then
111                        Range("A" & rw1.Row).Value = 10
                           If IsEmpty(Range("L" & rw1.Row).Value) Then
                                Range("L" & rw1.Row).Value = "Gestuurd op: " _
                                & Format(Now, "dd-mm-yyyy hh:mm")
                           Else
                                Range("L" & rw1.Row).Value = "Gestuurd op: " _
                                & Format(Now, "dd-mm-yyyy hh:mm") & " | " & Range("L" & rw1.Row).Value
                           End If
119                    End If
99                End Select
            End If
        Next rw1
89    End With

120 Sheets(Sht).Range("L" & rw.Row).Value = 1

79 End If
Volgende:
'Start clean, clear sjabloon
121 Sheets("EmailSjabloon").Select
122 Range("B2:B6, C2, L2, U1:W1000").Select
123 Selection.ClearContents
124 Selection.Borders.LineStyle = xlNone
125 Sheets(Sht).Range("L1").Value = ""

39 Next rw

EindeMailing:
130 Dim TotalMail As Integer
131 Dim SendedMail As Integer
132 Dim MailError As Integer

Sheets(Sht).Range("L1").FormulaR1C1 = "=SUM(R[1]C:R[" & EindData - 1 & "]C)"


133 TotalMail = Sheets(Sht).Range("M2", "M" & EindData).count
134 SendedMail = Sheets(Sht).Range("L1").Value
135 MailError = TotalMail - SendedMail
136 Sheets(Sht).Range("L1", "M" & EindData).ClearContents
137 CertNietBewerkbaar

138 Sheets("EmailSjabloon").Visible = xlSheetVeryHidden
139 Sheets(Sht).Visible = xlSheetVeryHidden

140 BackgroundFunction.AutoCloseMessage Tekst:=SendedMail & " mails are prepared" & vbNewLine & _
        MailError & " have given a error"
Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Function Mail_Range(Sht As String)
    
    Dim OutApp As Object
    Dim OutMail As Object
    Dim FindString As String
    Dim rng As Range
    Dim rw As Integer
    Dim clmn As Integer
    Dim Clm As Integer
    
SubName = "'Mail_Range'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
      
1    Sheets("EmailSjabloon").Visible = xlSheetVisible
2    Sheets("EmailSjabloon").Select
    
SelectLanguage:
        'Search for a the right range
3        FindString = Sheets("EmailSjabloon").Range("B5").Value
        
        'If there is a language selected, select right email template
10        If FindString <> "" Then
20        Select Case Sht
                
        Case "Aanvragen" 'nieuwe aanvraag maken
22        Clm = 1 'Kolom A

        Case "Email"
26        Clm = 11 'Kolom K
        
29        End Select
        
11        SearchRng = ColLett(Clm) & ":" & ColLett(Clm)
30        With Sheets("EmailSjabloon").Columns(SearchRng)
31            Set rng = .Find(What:=FindString, _
                            After:=.Cells(1), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
40            If Not rng Is Nothing Then
41                rw = rng.Row

50                   With application
                        .EnableEvents = False
                        .ScreenUpdating = False
59                    End With
                
                    Set OutApp = CreateObject("Outlook.Application")
                    Set OutMail = OutApp.CreateItem(0)
                
                    On Error Resume Next
60                    With OutMail
                        .To = Sheets("EmailSjabloon").Range("B3").Value
                        .CC = Sheets("EmailSjabloon").Range("B4").Value
                        .BCC = ""
                        
                        .Subject = GetSubject(rw, Clm)
                        .HTMLBody = GetBody(rw, Clm)
                        .Display   'or use .Send for direct send
69                    End With
                    On Error GoTo ErrorText:
                
70                    With application
                        .EnableEvents = True
                        .ScreenUpdating = True
79                    End With
          
          'Language is empty
45            Else
80                If FindString = "0" Then
                'MsgBox "Language is not valid. Value is '0'"
81                Sheets(Sht).Range("L1").Value = False
82                Exit Function
                
83                ElseIf FindString <> "EN" Then
                'MsgBox "Language '" & FindString & "' is not set. Language is set to English"
                
84                Sheets("EmailSjabloon").Range("B5").Value = "EN"
85                GoTo SelectLanguage:
                
86                Else
                'MsgBox "There is a problem with setting the standard language. This contact will be skiped"
87                Sheets(Sht).Range("L1").Value = False
88                Exit Function
89                End If
49            End If
39        End With
15        Else
16            BackgroundFunction.AutoCloseMessage Tekst:="There is a problem to get the language", Interval:=1
19        End If

    Set OutMail = Nothing
    Set OutApp = Nothing
    Sheets(Sht).Range("L1").Value = True
    
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Function GetSubject(rw As Integer, Coll As Integer)

SubName = "'GetSubject'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
    
1 Clm = ColLett(Coll + 1)
2 Clm2 = ColLett(Coll + 3)
   
3 GetSubject = ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 1).Value & " " & _
                ThisWorkbook.Sheets("EmailSjabloon").Range(Clm2 & rw + 1).Value

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Function

Function GetBody(rw As Integer, Coll As Integer)

    Dim StrBodyOpen As String 'opening text of the email
    Dim StrBodyClose As String 'end text of the email
    Dim rngHtml As Range 'Range for the changing body info (certificates)
   
SubName = "'GetBody'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

1 Clm = ColLett(Coll)
2 Clm2 = ColLett(Coll + 2)

3     Range(Clm & rw + 8, Clm2 & rw + 8).Copy
4    Range("U1").PasteSpecial (xlPasteValues)
    
5    Set rngHtml = Nothing
6    Eind = Range("U1000").End(xlUp).Row

8    Set tbl = Range("U1", "W" & Eind)
        'Tabel opmaak
9   Columns(tbl.Column).EntireColumn.AutoFit
    
    Range("U1:W1").Font.Bold = True
    
10    With tbl.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
         .Weight = xlThin
19    End With
    
20    With tbl.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
21    End With

    'opmaak certificaat informatie
    With Range("V1", "W" & Eind)
        .ColumnWidth = 13.57
        .NumberFormat = "d/m/yyyy"
    End With
    
    
30    StrBodyOpen = ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 3).Value & ThisWorkbook.Sheets("EmailSjabloon").Range("B" & rw + 3).Value & "<br>" & _
                    ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 4).Value & "<br>" & _
                    ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 6).Value

7    Set rngHtml = Sheets("EmailSjabloon").Range("U1", "W" & Eind).SpecialCells(xlCellTypeVisible)
   
31    StrBodyClose = "<br>" & ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 10).Value & "<br>" & _
                    ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 12).Value & "<br>" & _
                    ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 14).Value & "<br>" & _
                    ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 16).Value & "<br>" & _
                    ThisWorkbook.Sheets("EmailSjabloon").Range(Clm & rw + 18).Value & "<br>"
Maken:
    
40    GetBody = StrBodyOpen & RangetoHTML(rngHtml) & StrBodyClose
    
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

SubName = "'RangeToHtml'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
    
1    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy hh-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
2    rng.Copy
3    Set TempWB = Workbooks.Add(1)
10    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo ErrorText:
19    End With

    'Publish the sheet to a htm file
20    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
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

Exit Function
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Function
