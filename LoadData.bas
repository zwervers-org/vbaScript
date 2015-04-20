Attribute VB_Name = "LoadData"
Sub OpenFile()

SubName = "'OpenFile'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim FileSlct As String
Dim fDialog As Office.FileDialog
Dim FileLoad As Boolean
Dim wb As Workbook
Dim x As String
Dim y As String
Dim Ext As String

x = ActiveWorkbook.Name

CertBewerkbaar

With Workbooks(x).Sheets("Certificaten")
If Range("A1") <> "" Then
'Save old data
SaveOldData

SavePDF

'See what file will be selected
GoTo OpenFile:

Else

GoTo OpenFile:

End If
End With

'Clear existing data
ShtClear:


'Open a file as datafile
OpenFile:
    
Set fDialog = application.FileDialog(msoFileDialogFilePicker)


With application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Choose file to load (only TXT or XLS-files)"
    .Filters.Clear
    .Filters.Add "Text files", "*.txt"
    .Filters.Add "XLS files", "*.xls"

    If .Show = True Then
      If .SelectedItems.count = 1 Then
            application.ScreenUpdating = View("Updte")
            FileSlct = .SelectedItems(1)
            
            Ext = Right$(FileSlct, Len(FileSlct) - InStrRev(FileSlct, "."))
            
        Select Case (Ext)
            Case "txt"
                Workbooks.OpenText Filename:=FileSlct, Origin:= _
                xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                xlDoubleQuote, ConsecutiveDelimiter:=False, Semicolon:=True, _
                Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
                Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1)), TrailingMinusNumbers _
                :=True
                
            Case "xls"
                Workbooks.Open Filename:=FileSlct
                
            Case Else
                BackgroundFunction.AutoCloseMessage Tekst:="Can't find extention"
        End Select
        
        y = ActiveWorkbook.Name
        
    With Workbooks(y)
    Set CertRng = Range("C1", Range("C1000").End(xlUp))
        Select Case Range("A1").Value
            Case "OTC-Holland"
            'Clear sheet first
                Workbooks(x).Activate
                ClearSheet ("Certificaten")
            'Start new certificate entery:
                LoadNewData y, "NL", x, Ext
                GoTo Einde:
            
        'Extention data
            Case "OTC-USA"
                Workbooks(x).Activate
                For Each rw In CertRng
                    If Range("B" & rw.Row).Value = "US" Then
                    BackgroundFunction.AutoCloseMessage Tekst:="USA Data already loaded"
                    GoTo Einde:
                End If
                Next rw
                
                LoadNewData y, "US", x, Ext
                GoTo Einde:
            
            Case "OTC-Belgium bvba"
                Workbooks(x).Activate
                For Each rw In CertRng
                    If Range("B" & rw.Row).Value = "BE" Then
                    BackgroundFunction.AutoCloseMessage Tekst:="Belgium Data already loaded"
                    GoTo Einde:
                End If
                Next rw
            
                LoadNewData y, "BE", x, Ext
                GoTo Einde:
            
            Case "Flevo Fresh B.V."
                Workbooks(x).Activate
                For Each rw In CertRng
                    If Range("B" & rw.Row).Value = "FF" Then
                    BackgroundFunction.AutoCloseMessage Tekst:="FlevoFresh Data already loaded"
                    GoTo Einde:
                End If
                Next rw
                
                LoadNewData y, "FF", x, Ext
                GoTo Einde:
        
        'Not recognized file
            Case Else
                MsgBox "Importsystem does not recognize datafile"
                GoTo Einde:
            End Select
    End With
    
    Else
         MsgBox "You can only select ONE file."
         
         GoTo OpenFile:
    End If
      
    Else
      MsgBox "You canceled the file loading."
      
      LoadOldData (Format(Range("A1").Value, "mm-dd-yyyy"))
      GoTo Einde:
    End If
End With
Exit Sub

Einde:
Workbooks(x).Activate
If y <> "" Then
    Workbooks(y).Close
    
    RemoveSourceFile = MsgBox("Do you want to remove the sourcefile?", vbYesNo, "Remove source file?")
    
    If RemoveSourceFile = vbYes Then
        SetAttr FileSlct, vbNormal
        Kill FileSlct
        Error.DebugTekst Tekst:="Removed source file: " & FileSlct, FunctionName:=SubName
    End If
End If

If Workbooks(x).Sheets("Certificaten").ProtectContents <> True Then CertNietBewerkbaar

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next
GoTo Einde
    
End Sub

Function LoadNewData(DataWb As String, DataBV As String, SourceFile As String, Extention As String)

SubName = "'LoadNewData'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim StrName As String
Dim rng As Range
Dim rw As Range
Dim rij As Integer
'Dim OldDate As Date
'Dim NewDate As Date

jaar = Year(Now)
'Name of the base file
NmWb = SourceFile

'Extention of the datafile
Ext = Extention

'Prepare the datafile
Workbooks(DataWb).Activate
With Workbooks(DataWb)

Select Case (Ext)
    Case "txt"
    
    'Datum in juiste cel plaatsen
    If Range("A2") = Range("A1") Then
        Range("B3").Copy Destination:=Range("C2")
    ElseIf Range("A2") = "Zakenrelaties (Certificaten)" Then
        Range("B2").Copy Destination:=Range("C2")
    Else
        StrName = InputBox("Problem to get de date of the file." & vbNewLine & _
        "Please give the date you checked the validity till: (mm-dd-yyyy)", "Give validity date")
        Range("C2").Value = StrName
    End If
    
    NewDate = Range("C2").Value
    
    NewDate = Format(NewDate, "mm-dd-yyyy")
    Error.DebugTekst ("NewDate= " & NewDate)

OpnieuwTxt:
        Einde = Range("A1000").End(xlUp).Row
        
        Set rng = Range("A2", "L" & Einde)
        
        For Each rw In rng.Rows
            If rw.Columns(5) = "" Then
                If rw.Row > 2 Then rw.Delete
            ElseIf rw.Columns(1) = "© 2001-" & jaar & " Neddox ®" Then
                rw.Delete
            ElseIf rw.Columns(1) = "Code" Then
                rw.Delete
            ElseIf rw.Columns(5) = Range("E" & rw.Row - 1) Then
                If Range("D" & rw.Row) = Range("D" & rw.Row - 1) Then
                    If Range("C" & rw.Row) = Range("C" & rw.Row - 1) Then
                        If Range("A" & rw.Row) = Range("A" & rw.Row - 1) Then
                            rw.Delete
                        End If
                    End If
                End If
            End If
        Next rw
        
        'Wanneer er nog cellen bestaan overnieuw alles bekijken en verwijderen.
        For Each rw In rng.Rows
            If rw.Columns(5) = "" Then
                If rw.Row > 2 Then GoTo OpnieuwTxt:
            ElseIf rw.Columns(1) = "© 2001-" & jaar & " Neddox ®" Then
                GoTo OpnieuwTxt:
            ElseIf rw.Columns(1) = "Code" Then
                GoTo OpnieuwTxt:
            ElseIf rw.Columns(5) = Range("E" & rw.Row - 1) Then
                If Range("D" & rw.Row) = Range("D" & rw.Row - 1) Then
                    If Range("C" & rw.Row) = Range("C" & rw.Row - 1) Then
                        If Range("A" & rw.Row) = Range("A" & rw.Row - 1) Then
                            GoTo OpnieuwTxt:
                        End If
                    End If
                End If
            End If
        Next rw
        
        'Onnodige informatie verwijderen
            Columns("F:L").Delete Shift:=xlToLeft
                
    Case "xls"
        ActiveSheet.Shapes.SelectAll
        Selection.Delete
        
        'Wanneer er nog cellen bestaan overniew alles bekijken.
OpnieuwXls:
        Einde = Range("A1000").End(xlUp).Row
        
        Set rng = Range("A1", "L" & Einde)
        
        For Each rw In rng.Rows
            If rw.Columns(1) = "" Then
                rw.Delete
            ElseIf rw.Columns(1) = "© 2001-" & jaar & " Neddox ®" Then
                rw.Delete
            ElseIf rw.Columns(1) = "Code" Then
                rw.Delete
                End If
        Next rw
        
        'Wanneer er nog cellen bestaan overnieuw alles bekijken en verwijderen.
        For Each rw In rng.Rows
            If rw.Columns(1) = "" Then
                GoTo OpnieuwXls:
            ElseIf rw.Columns(1) = "© 2001-" & jaar & " Neddox ®" Then
                GoTo OpnieuwXls:
            ElseIf rw.Columns(1) = "Code" Then
                GoTo OpnieuwXls:
                End If
        Next rw
        
        'Onnodige informatie verwijderen
            Columns("F:L").Delete Shift:=xlToLeft

    Case Else
        MsgBox "Importsystem does not recognize datafile"
        GoTo Einde:
    End Select
End With

'delete password protection
Workbooks(NmWb).Activate
CertBewerkbaar

'Set Copy-range
With Workbooks(DataWb)
    .Activate
    With .Sheets(ActiveSheet.Name)
        Einde = .Range("A1000").End(xlUp).Row
    
        If .Range("A3").Value = "Code" Then
            Set rng = .Range("A4", "E" & Einde)
        ElseIf .Range("A2").Value = "Zakenrelaties (Certificaten)" Then
            Set rng = .Range("A3", "E" & Einde)
        ElseIf .Range("C2").Value = NewDate Then
            Set rng = .Range("A3", "E" & Einde)
        Else
            Set rng = .Range("A2", "E" & Einde)
        End If
    End With
End With

Select Case DataBV
Case "NL"
    GoTo CopyNwData:

Case "US"
    GoTo CopyExtrData:

Case "BE"
    GoTo CopyExtrData:

Case "FF"
    GoTo CopyExtrData:

End Select

Exit Function

CopyNwData:

    With Workbooks(DataWb)
        .Activate
        rng.Copy
    End With

'Paste existing data
    With Workbooks(NmWb)
        .Activate
        With .Sheets("Certificaten")
            .Select
            .Range("C2").PasteSpecial xlPasteValues

        'Past date of new data
            .Range("A1").Value = NewDate
            .Range("A1").Select
        End With
    End With
BackgroundFunction.AutoCloseMessage Tekst:="New data loaded. Checked for validity till: " & Range("A1").Value

GoTo SorteerActie

CopyExtrData:
    'Check date of the files
    With Workbooks(NmWb)
        .Activate
        
        OldDate = .Sheets("Certificaten").Range("A1").Value
    
        OldDate = Format(OldDate, "mm-dd-yyyy")
        Error.DebugTekst ("OldDate= " & OldDate)
    End With
    
    If OldDate = NewDate Then
    
    'Copy extra data
    With Workbooks(DataWb)
        .Activate
            rng.Copy
    End With
    
    'Paste existing data
    With Workbooks(NmWb)
        .Activate
        With .Sheets("Certificaten")
            .Select
            
            Dim EntryRng As Range
            
            EindeNmWb = .Range("C1").End(xlDown).Row + 1
            
            .Range("C" & EindeNmWb).PasteSpecial xlPasteValues
            
            Set EntryRng = .Range("C" & EindeNmWb, "C" & .Range("C1").End(xlDown).Row)
            
            
        'Paste the division code before the pasted enteries
            For Each rw In EntryRng
                .Range("B" & rw.Row).Value = DataBV
            Next rw
            
            'Sort on Code
            EindeSort = .Range("C1000").End(xlUp).Row
            .Columns("A:P").Select
            .Sort.SortFields.Clear
            .Sort.SortFields.Add Key:=.Range( _
                "C1", "C" & EindeSort), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
            With .Sort
                .SetRange Range("A1", "P" & EindeSort)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End With
    End With
    
Opnieuw:
    'delete password protection
    With Workbooks(NmWb)
        .Activate
        CertBewerkbaar
        With .Sheets("Certificaten")
            DataRng = .Range("C1000").End(xlUp).Row
            For rij = 1 To DataRng Step 1
                If .Range("B" & rij).Value = DataBV Then
                    If .Range("C" & rij).Value = Range("C" & rij - 1).Value Then
                        If .Range("E" & rij).Value = Range("E" & rij - 1).Value Then
                            If .Range("F" & rij).Value = Range("F" & rij - 1).Value Then
                                If .Range("G" & rij).Value = Range("G" & rij - 1).Value Then
                                    .Range("B" & rij - 1).Value = "+" & DataBV
                                    .Rows(rij).Delete (Shift = xlUp)
                                    rij = rij - 2
                                    
                                Else
                                Dim Answer As String
                                Dim MyNote As String
                                
                                    'Place your text here
                                    MyNote = "For supplier: " & .Range("D" & rij).Value & Chr(10) _
                                                & "Are ` " & .Range("G" & rij - 1).Value & " ` and ` " _
                                                & .Range("G" & rij).Value & " ` the same certificates?"
                                
                                    'Display MessageBox
                                    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Dobble certificates?")
                                
                                    If Answer = vbYes Then
                                        .Range("B" & rij - 1).Value = "+" & DataBV
                                        .Rows(rij).Delete (Shift = xlUp)
                                        rij = rij - 2
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next rij
    
            Range("A1").Select
    
        End With
        'set password protection
        CertNietBewerkbaar
    End With
    
    BackgroundFunction.AutoCloseMessage Tekst:="Data extention: " & DataBV & " loaded"
    
    Else
    MsgBox "Can't load data, date from source file and base file are not the same:" & Chr$(10) _
            & "Basefile date: " & Range("C1").Value & Chr$(10) _
            & "New Sourcefile date: " & Range("C2").Value
    
    End If

SorteerActie:
'Sort on Code
    With Workbooks(NmWb)
        .Activate
        With .Sheets("Certificaten")
            EindeSort = .Range("C1000").End(xlUp).Row
            Columns("A:P").Select
            .Sort.SortFields.Clear
            .Sort.SortFields.Add Key:=.Range( _
                 "C1", "C" & EindeSort), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                 xlSortNormal
            With .Sort
                .SetRange Range("A1", "P" & EindeSort)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End With
    End With

Einde:
With Workbooks(NmWb)
    .Activate
    With .Sheets("Certificaten")
        .Range("A1").Select
    End With
End With

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function

Function LoadOldData(GotoDate As String)

SubName = "'LoadOldData'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim Sht As Worksheet
Dim rng As Range
Dim Nm As String

CertBewerkbaar

If GotoDate = "" Then
Range("AC1").Value = "Large"
LoadDate.Show

    If IsNull(LoadDate.FindString.Value) Then
        BackgroundFunction.AutoCloseMessage Tekst:="The process is canceled.", Titel:="Task is canceled"
        CertNietBewerkbaar
        Exit Function
    Else
        GotoDate = LoadDate.FindString.Value
    End If

End If

If Range("A1") <> "" Then 'Wanneer er informatie instaat eerst opslaan en sheet legen

    'Save old data
    SaveOldData
    
    'Clear existing data
    ClearSheet ("Certificaten")
    
    GoTo LoadDataBack:

Else

    If Range("D2") <> "" Then
    MsgBox "There is information in the file" & Chr$(10) _
            & "Please check the data in 'Certificaten' and clean the sheet" & Chr$(10) _
            & "Process is canceled"
            
            Exit Function
    Else
    GoTo LoadDataBack:
    End If

Exit Function

LoadDataBack:
    
        Sheets(GotoDate).Visible = xlSheetVisible
        Sheets(GotoDate).Select
        GoTo CopyData:

End If
Exit Function

CopyData:
'Set workmodus
    Sheets("Certificaten").Select
    CertBewerkbaar

    Sheets(GotoDate).Visible = xlSheetVisible
    Sheets(GotoDate).Select

'Copy existing data

    Einde = Range("C1000").End(xlUp).Row

    Set rng = Range("A2", "G" & Einde)
    
    rng.Copy

'Paste existing data
    Sheets("Certificaten").Select
    Range("A2").PasteSpecial xlPasteValues

'Second copy existing data
    Sheets(GotoDate).Select
    
    Set rng = Range("L2", "L" & Einde)
    
    rng.Copy
    
'Second paste existing data
    Sheets("Certificaten").Select
    
    Range("L2").PasteSpecial xlPasteValues

'Copy date back
    Sheets(GotoDate).Select
    
    Range("A1").Copy
    
'Copy date back
    Sheets("Certificaten").Select
    
    Range("A1").PasteSpecial xlPasteValues
    
'Back to workable sheet
    Range("A2").Select
    CertNietBewerkbaar
    
    Sheets(GotoDate).Visible = xlSheetVerryHidden

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Function
