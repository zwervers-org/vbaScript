Attribute VB_Name = "RelHistory"
Sub SearchHist()
Attribute SearchHist.VB_ProcData.VB_Invoke_Func = "F\n14"

If Sheets("Certificaten") Is ActiveSheet Then

HistRel.Show

End If

End Sub

Function SearchHistory(code As String, SearchType As String)

SubName = "'ShearchHistory'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim StrOpm() As String 'string for comments
Dim StrCert() As String 'string for certificate information
Dim StrToDate() As String 'string for expiering date
Dim StrAction() As String 'string for the action token
Dim StrDiv() As String 'string for the division
Dim StrName() As String 'string for company name
Dim StrCode() As String  'string for Neddox code
Dim ShtDate As String 'search date
Dim StrCnt As Boolean 'counter
Dim StrYes As Boolean
Dim rng As Range
Dim StrNr As Integer
Dim CodeList() As String
                    
Dim txt As String
Dim k As Integer

'clear date buffer
With Sheets("Certificaten")
    .Select
    CertBewerkbaar
    Range("S1", "AA" & Range("S2").End(xlDown).Row).ClearContents
    CertNietBewerkbaar
End With

Error.DebugTekst Tekst:="Search with> code: " & code & " | SearchType: " & SearchType, FunctionName:=SubName

If SearchType <> "Neddox" Then
'zien of er aliassen zijn opgegeven voor het certificaat
34    FindString = code
    
    
40    With Sheets("DATA").Columns(26)
41    Dim rng1 As Range
42    Dim rng2 As Range
    
43    Set rng1 = Nothing
45                Set rng1 = .Find(What:=FindString, _
                                After:=.Cells(1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)

50        If Not rng1 Is Nothing Then
            'Certificaat is gevonden
            With Sheets("DATA")
                LastColumn = BackgroundFunction.ColLett(.Range("Z" & rng1.Row).End(xlToRight).Column)
                
                If .Range("AA" & rng1.Row).Value <> "" Then
                    Set rng2 = .Range("AA" & rng1.Row & ":" & LastColumn & rng1.Row)
                    
                    ReDim CodeList(0 To rng2.Columns.count) As String
                    
60                  For AliasNr = 0 To rng2.Columns.count - 1
                        AliasColumn = BackgroundFunction.ColLett(AliasNr + 27)
                        CodeList(AliasNr) = .Range(AliasColumn & rng1.Row).Value
                    Next AliasNr
                Else
                    ReDim CodeList(0 To 1) As String
                    
                    CodeList(0) = code
                End If
            End With
          
          Else
            ReDim CodeList(0 To 1) As String
            
            CodeList(0) = code
          
          End If
        
49    End With
End If

'show all archived data sheets
For i = 1 To Sheets.count
  If InArray("Sheets", Sheets(i).Name) Then
    'skip this sheet
  Else
    ShowOneSheet (Sheets(i).Name)
    
  'there is no array yet
    StrCnt = False
    StrYes = False
    StrNr = 0


    Select Case SearchType
        Case "Neddox"
            With Sheets(i) 'look if the code is in the sheet
            
                Einde = Range("C1000").End(xlUp).Row
                
                Set rng = Range("C2", "C" & Einde)
                For Each rw In rng
                    If Range("C" & rw.Row).Value = code Then
                            StrYes = True
                      'make next array entries
                        If StrCnt = True Then
                            ReDim Preserve StrOpm(0 To UBound(StrOpm) + 1) As String
                            ReDim Preserve StrCert(0 To UBound(StrCert) + 1) As String
                            ReDim Preserve StrToDate(0 To UBound(StrToDate) + 1) As String
                            ReDim Preserve StrAction(0 To UBound(StrAction) + 1) As String
                            ReDim Preserve StrDiv(0 To UBound(StrAction) + 1) As String
                            ReDim Preserve StrName(0 To UBound(StrAction) + 1) As String
                            ReDim Preserve StrCode(0 To UBound(StrAction) + 1) As String
                            
                            StrNr = StrNr + 1
                      'make first array entry
                        Else
                            ReDim StrOpm(0 To 0) As String
                            ReDim StrCert(0 To 0) As String
                            ReDim StrToDate(0 To 0) As String
                            ReDim StrAction(0 To 0) As String
                            ReDim StrDiv(0 To 0) As String
                            ReDim StrName(0 To 0) As String
                            ReDim StrCode(0 To 0) As String
                            
                            ShtDate = ActiveSheet.Name
                            
                            StrCnt = True
                        End If
        
                      'put data into array
                        StrAction(StrNr) = .Range("A" & rw.Row).Value
                        StrDiv(StrNr) = .Range("B" & rw.Row).Value
                        StrCode(StrNr) = .Range("C" & rw.Row).Value
                        StrName(StrNr) = .Range("D" & rw.Row).Value
                        StrToDate(StrNr) = .Range("F" & rw.Row).Value
                        StrCert(StrNr) = .Range("G" & rw.Row).Value
                        StrOpm(StrNr) = .Range("L" & rw.Row).Value
                        
                    End If
                Next rw
            End With
            GoTo PlaceInBuffer
            
        Case "Certificate"
            With Sheets(i) 'look if the code is in the sheet
                
                Einde = .Range("C1000").End(xlUp).Row
                
                Set rng = .Range("G2", "G" & Einde)
                For Each rw In rng
                  For CodeInList = 0 To UBound(CodeList) - 1
                    If InStr(1, .Range("G" & rw.Row).Value, CodeList(CodeInList), 1) > 0 Then
                        StrYes = True
                      'make next array entries
                        If StrCnt = True Then
                            ReDim Preserve StrOpm(0 To UBound(StrOpm) + 1) As String
                            ReDim Preserve StrCert(0 To UBound(StrCert) + 1) As String
                            ReDim Preserve StrToDate(0 To UBound(StrToDate) + 1) As String
                            ReDim Preserve StrAction(0 To UBound(StrAction) + 1) As String
                            ReDim Preserve StrDiv(0 To UBound(StrAction) + 1) As String
                            ReDim Preserve StrName(0 To UBound(StrAction) + 1) As String
                            ReDim Preserve StrCode(0 To UBound(StrAction) + 1) As String
                            
                            StrNr = StrNr + 1
                      'make first array entry
                        Else
                            ReDim StrOpm(0 To 0) As String
                            ReDim StrCert(0 To 0) As String
                            ReDim StrToDate(0 To 0) As String
                            ReDim StrAction(0 To 0) As String
                            ReDim StrDiv(0 To 0) As String
                            ReDim StrName(0 To 0) As String
                            ReDim StrCode(0 To 0) As String
                            
                            ShtDate = ActiveSheet.Name
                            
                            StrCnt = True
                        End If
        
                      'put data into array
                        StrAction(StrNr) = .Range("A" & rw.Row).Value
                        StrDiv(StrNr) = .Range("B" & rw.Row).Value
                        StrCode(StrNr) = .Range("C" & rw.Row).Value
                        StrName(StrNr) = .Range("D" & rw.Row).Value
                        StrToDate(StrNr) = .Range("F" & rw.Row).Value
                        StrCert(StrNr) = .Range("G" & rw.Row).Value
                        StrOpm(StrNr) = .Range("L" & rw.Row).Value
                    End If
                  Next CodeInList
                Next rw
            End With
            GoTo PlaceInBuffer
            
        Case Else
            MsgBox "Program does not understand where to search for" & vbNewLine _
                    & "Program does hard stop"
            RelHistory.CloseOverview ("No")
            End
    End Select

PlaceInBuffer:
    Error.DebugTekst Tekst:="Place search data in buffer", FunctionName:=SubName
    
            HideOneSheet (Sheets(i).Name)
                If StrYes = True Then
                    With Sheets("Certificaten")
                    CertBewerkbaar
                        Eind = .Range("U10000").End(xlUp).Row
                        For strpos = 0 To UBound(StrCert)
                            .Range("S" & Eind + strpos + 1).Value = Eind + strpos + 1
                            'Range("T" & Eind + StrPos + 1).Value = ShtDate
                            If ShtDate <> "" Then 'Zet SheetName (datum) om in text
                                SplShtDate = Split(ShtDate, "-")
                                ToDay = SplShtDate(1)
                                ToMonth = SplShtDate(0)
                                ToYear = SplShtDate(2)
                                .Range("T" & Eind + strpos + 1).Value = ToDay & "-" & ToMonth & "-" & ToYear
                            End If
                            
                            .Range("U" & Eind + strpos + 1).Value = StrCode(strpos) 'Leveranciers Neddox Code
                            
                            .Range("V" & Eind + strpos + 1).Value = StrName(strpos) 'Naam leverancier
                            
                            If StrDiv(strpos) = "" Then 'Wanneer geen divisie dan divisie is NL
                                    .Range("W" & Eind + strpos + 1).Value = "NL" 'alleen een NL divisie
                            Else
                                If InStr(1, StrDiv(strpos), "+", 1) > 0 Then 'er is naast een NL divisie ook een andere divisie
                                    .Range("W" & Eind + strpos + 1).Value = "NL" & StrDiv(strpos)
                                Else
                                    .Range("W" & Eind + strpos + 1).Value = StrDiv(strpos) 'alleen een andere divisie
                                End If
                            End If
                            
                             If StrAction(strpos) = "" Then 'wanneer er een actie is ingevoerd de actie opzoeken adhv de actiecode
                                StrAction(strpos) = 0
                             End If
                            .Range("X" & Eind + strpos + 1).FormulaR1C1 = _
                            "=IF(" & StrAction(strpos) & " =0,DATA!R3C11, " _
                            & "VLOOKUP(" & StrAction(strpos) & ",AfwerkCodes,2))"
                            
                            If StrToDate(strpos) <> "" Then 'Zet geldig tot datum omzetten in text
                                SplStrToDate = Split(StrToDate(strpos), "-")
                                ToDay = SplStrToDate(1)
                                ToMonth = SplStrToDate(0)
                                ToYear = SplStrToDate(2)
                                .Range("Y" & Eind + strpos + 1).Value = ToDay & "-" & ToMonth & "-" & ToYear
                            End If
                            
                            .Range("Z" & Eind + strpos + 1).Value = StrCert(strpos) 'certificaat
                            
                            .Range("AA" & Eind + strpos + 1).Value = StrOpm(strpos) 'opmerking
                        Next strpos
                    CertNietBewerkbaar
                    End With
                End If
                
              'Erase arrays
                Erase StrOpm
                Erase StrCert
                Erase StrToDate
                Erase StrAction
                ShtDate = ""
          End If
    Next i

CertBewerkbaar
Einde = Range("T10000").End(xlUp).Row
    ActiveWorkbook.Worksheets("Certificaten").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Certificaten").Sort.SortFields.Add Key:=Range( _
        "T2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Certificaten").Sort
        .SetRange Range("T2:AA" & Einde)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    CertNietBewerkbaar

CertBewerkbaar
    With Sheets("Certificaten")
        .Range("T1").FormulaR1C1 = "Sheet"
        .Range("U1").FormulaR1C1 = "Code"
        .Range("V1").FormulaR1C1 = "Naam"
        .Range("W1").FormulaR1C1 = "Div"
        .Range("X1").FormulaR1C1 = "Actie"
        .Range("Y1").FormulaR1C1 = "Tot Datum"
        .Range("Z1").FormulaR1C1 = "Certificaat"
        .Range("AA1").FormulaR1C1 = "Opmerking"
    End With

CertNietBewerkbaar

HistorySearch.Show

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function

Function PublishHistorySearch()

SubName = "'PublishHistorySearch'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function
Einde = Range("T10000").End(xlUp).Row

qv = "Quick view data?"

TextBox = False

Error.DebugTekst Tekst:="View history search", FunctionName:=SubName

If Einde >= 2 Then
    If TextBox = False Then
        Dim ValueRange As Range
        
        Set ValueRange = Range("T2", "AA" & Einde)
        
        'put the values into a userform textbox
        
            'For rw = 2 To Einde
            '    With HistorySearch.ResultSearch
            '        .Value = Application.WorksheetFunction.VLookup(Range("S" & rw), Sheets("Certificaten").Range("T" & rw, "Z" & rw), 1, False)
            '       .Value = rw.Offset(0, 1)
            '    End With
            'Next rw
            
            'HistorySearch.ResultSearch.AutoSize
        
        HistorySearch.ListBox1.Clear
        
        HistorySearch.ListBox1.ColumnCount = ValueRange.Columns.count
        
        HistorySearch.ListBox1.RowSource = ValueRange.Address
        HistorySearch.ListBox1.ColumnHeads = True
        HistorySearch.ListBox1.TextColumn = -1
        
        GoTo Skip
        For Each c In ValueRange.Columns(1).SpecialCells(xlCellTypeVisible)
            With HistorySearch.ListBox1
                .AddItem c.Value
                For colmn = 1 To ValueRange.Columns.count
                    .List(.ListCount - 1, colmn) = c.Offset(0, colmn).Value
                    ValueRange.Columns(colmn).AutoFit
                Next colmn
            End With
        Next c
                
Skip:
        'HistorySearch.ListBox1.ColumnWidths = ValueRange.Columns(1).Width

    Else
        Dim myData
        
        Set rng = Range("U1", "AA" & Einde)
        
        myData = rng.Value
        
        'Put the values into a messagebox
        For k = 1 To UBound(myData, 1)
            txt = txt & vbNewLine _
                            & Cells(k, 20).Value & vbTab & Cells(k, 21).Value & vbTab & Cells(k, 22).Value _
                            & vbTab & Cells(k, 23).Value & vbTab & Cells(k, 24).Value & vbTab _
                            & Cells(k, 25).Value & vbTab & Cells(k, 26).Value
        
        Next k
        
        'txt = txt & vbNewLine & vbNewLine & vbTab & "DO YOU WHANT TO SAVE THIS?"
        'CloseOverview (txt)
        
        CloseOverview ("No") 'Quickview activated
        
    End If
    
    Else
        RelHistory.CloseOverview (False)
        HistorySearch.Hide
        HistRel.Show
End If

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function

Function CloseOverview(txt As String)

SubName = "'CloseOverview'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Select Case txt 'save data?
    Case "No"
        Response = 7
    Case "Yes"
        Response = 6
    Case False
        Response = MsgBox("There are no values for this code: " & code & " in the history" & vbNewLine _
                & "Try it also with kapitals", vbOKOnly + vbCritical + vbDefaultButton1, Title:="NO HISTORY")
    Case Is <> ""
        Response = MsgBox(txt, vbYesNo + vbInformation + vbDefaultButton2, Title:="Save history?")
    Case Else
        BackgroundFunction.AutoCloseMessage Tekst:="ERROR, no message to display!!!!"
        End
End Select

With Sheets("Certificaten")
    Einde = .Range("T10000").End(xlUp).Row
    
    Select Case Response
        Case 6
            .Range("T2").Select
        Case 7 Or 1
            CertBewerkbaar
            .Range("S1", "AA" & Einde).ClearContents
            CertNietBewerkbaar
            .Range("A2").Select
        Case Else
            CertBewerkbaar
            .Range("S1", "AA" & Einde).ClearContents
            CertNietBewerkbaar
            .Range("A2").Select
    End Select
End With

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function

Sub FillActions()
SubName = "'FillActions'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim StrOpm() As String
Dim StrCert() As String
Dim StrToDate() As String
Dim StrAction() As String
Dim StrDiv() As String
Dim ShtDate As String
Dim StrYes As Boolean
Dim rng As Range
                    
Dim txt As String
Dim k As Integer

LoadDate.Show

GotoDate = LoadDate.FindString.Value

Error.DebugTekst Tekst:="Fill data from: " & GotoDate & " > into sheet: " & Format(Range("A1").Value, "mm-dd-yyyy"), FunctionName:=SubName

If IsNull(GotoDate) Then
BackgroundFunction.AutoCloseMessage Tekst:="The process is canceled.", Titel:="Task is canceled"
Exit Sub

Else
Sheets("DATA").Range("X1").Value = "True" 'Let system know this is a automatic function that is loading data

BackgroundFunction.CertBewerkbaar
Sheets(GotoDate).Visible = xlSheetVisible

Einde = Sheets("Certificaten").Range("C1000").End(xlUp).Row
Set RngCert = Sheets("Certificaten").Range("C2", "C" & Einde)

    For Each rwCert In RngCert
    
    CodeStr = Sheets("Certificaten").Range("C" & rwCert.Row).Value
    CertStr = Sheets("Certificaten").Range("G" & rwCert.Row).Value
    VanDateStr = Sheets("Certificaten").Range("E" & rwCert.Row).Value
    TotDateStr = Sheets("Certificaten").Range("F" & rwCert.Row).Value
    
      'look if the code is in the sheet
        Einde = Sheets(GotoDate).Range("C1000").End(xlUp).Row
            
        Set rng = Sheets(GotoDate).Range("C2", "C" & Einde)
        For Each rw In rng
            If Sheets(GotoDate).Range("C" & rw.Row).Value = CodeStr Then
                If Sheets(GotoDate).Range("G" & rw.Row).Value = CertStr Then
                    If Sheets(GotoDate).Range("E" & rw.Row).Value = VanDateStr Then
                        If Sheets(GotoDate).Range("F" & rw.Row).Value = TotDateStr Then
                            If Sheets(GotoDate).Range("A" & rw.Row).Value <> "6" Then 'Check if status was 'ready'
                                Sheets(GotoDate).Range("A" & rw.Row).Copy
                                    Sheets("Certificaten").Range("A" & rwCert.Row).PasteSpecial xlPasteValues
                                
                                Sheets(GotoDate).Range("L" & rw.Row).Copy
                                    Sheets("Certificaten").Range("L" & rwCert.Row).PasteSpecial xlPasteValues
                                    
                            Else
                                Sheets("Certificaten").Range("A" & rwCert.Row).Value = "7"
                                
                                Sheets(GotoDate).Range("L" & rw.Row).Copy
                                    Sheets("Certificaten").Range("L" & rwCert.Row).PasteSpecial xlPasteValues
                                    Sheets("Certificaten").Range("L" & rwCert.Row).Value = "AFGEWERKT " & Left(Sheets(GotoDate).Name, 2) & "/" _
                                    & Mid(Sheets(GotoDate).Name, 4, 2) & " | " & Sheets("Certificaten").Range("L" & rwCert.Row).Value
                                
                            End If
                            
                        'Check if the Division is changed in time
                            If Sheets(GotoDate).Range("B" & rw.Row).Value = Sheets("Certificaten").Range("B" & rwCert.Row).Value Then
                                Sheets(GotoDate).Range("B" & rw.Row).Copy
                                    Sheets("Certificaten").Range("B" & rwCert.Row).PasteSpecial xlPasteValues
                            Else
                                'check if there is a + in the line / or missing
                                Plus = InStr(Sheets("Certificaten").Range("B" & rwCert.Row).Value, "+")
                                
                                    If Plus > 0 Then
                                        Sheets(GotoDate).Range("B" & rw.Row).Copy
                                            Sheets("Certificaten").Range("B" & rwCert.Row).PasteSpecial xlPasteValues
                                    Else
                                        Sheets("Certificaten").Range("A" & rwCert.Row).Value = "7"
                                        
                                        If IsEmpty(Sheets("Certificaten").Range("B" & rwCert.Row).Value) Then
                                            Sheets("Certificaten").Range("L" & rwCert.Row) = "Ook voor: NL" _
                                                & " | " & Sheets("Certificaten").Range("L" & rwCert.Row).Value
                                        Else
                                            Sheets("Certificaten").Range("L" & rwCert.Row) = "Ook voor: " _
                                                & Sheets("Certificaten").Range("B" & rwCert.Row).Value _
                                                & " | " & Sheets("Certificaten").Range("L" & rwCert.Row).Value
                                        End If
                                    End If
                            End If
                        End If
                    End If
                End If
            End If
        Next rw
    Next rwCert

Sheets("DATA").Range("X1").Value = "False" 'Let system know this was a automatic function that is loading data

Error.DebugTekst Tekst:="Data filled", FunctionName:=SubName

BackgroundFunction.CertNietBewerkbaar
End If

BackgroundFunction.AutoCloseMessage Tekst:="Actions form " & GotoDate & " is filled into the active overview."

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Sub

