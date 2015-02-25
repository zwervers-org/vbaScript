VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

'Clear BUGFIX Indicator
ActiveSht = ActiveSheet.Name

If ActiveSht <> "DATA" Then Sheets("DATA").Visible = xlSheetVisible
    
    With Sheets("DATA")
        .Select
        .Range("T21").ClearContents
        .Range("T22").ClearContents
        .Range("T23").ClearContents
        .Range("T24").ClearContents
        .Range("T25").ClearContents
        .Range("T26").ClearContents
    End With

Admin.HideAllSheets

End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)

If Target.Cells.count <> 1 Then
    Exit Sub
ElseIf IsError(Target.Value) Then
    Exit Sub
ElseIf ActiveSheet.Name <> "Certificaten" Then
    waarde = Target.Value
    
    Select Case waarde
    Case "m"
            Target.ClearContents
            
            MenuShow
    Case "b"
            Target.ClearContents
            
            ShowOneSheet ("Certificaten")
    End Select
End If

If ActiveSheet.Name <> "Certificaten" Then
    Exit Sub
Else
    If Target.Value = "d" Then 'Naar data sheet en rij met juiste Neddox-code selecteren
        Dim Zoekcode As String
        Dim FindZoekcode As Range
        
        Target.ClearContents 'cel weer leegmaken
        Rows(Selection.Row - 1).EntireRow.Select 'weer terug naar de vorige cel
        
        Zoekcode = Range("C" & Target.Row).Value 'Neddox code opslaan om te gebruiken in data sheet
        
        ShowOneSheet ("DATA") 'naar data sheet gaan
        'Range("A1").Select 'bovenaan beginnen
        
        On Error Resume Next 'voorkomen dat er een foutmelding komt wanneer er geen code gevonden wordt.
        
        With Sheets("DATA").Range("A:A")
            Set FindZoekcode = .Find(What:=Zoekcode, _
                                After:=.Cells(.Cells.count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False, _
                                SearchFormat:=False) 'Neddox code zoeken
            
            If Not FindZoekcode Is Nothing Then 'kijken of de waarde al bestaat
                Rows(FindZoekcode.Row).EntireRow.Select
                
            Else 'zoekcode bestaat nog niet
                LaatsteRij = Range("A1").End(xlDown).Row
                Range("A" & LaatsteRij + 1).Value = Zoekcode
                
                Rows(LaatsteRij + 1).EntireRow.Select 'Selecteer rij
            End If
        End With

'----------------
    ElseIf Target.Column = 1 Then 'er wordt een actie ingevoerd
        If IsNumeric(Target.Value) = False Then Exit Sub
                
        AutoCheck = Sheets("DATA").Range("X1").Value

        If AutoCheck = False Then
            '-----------------------De actie beschrijving opzoeken
            Dim FindString As String
            Dim RngString As Range
            FindString = Target.Value
                With Sheets("DATA").Range("J:J")
                    Set RngString = .Find(What:=FindString, _
                                    After:=.Cells(.Cells.count), _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
                End With
                
            If Not RngString Is Nothing Then
                ActionName = Sheets("DATA").Range("K" & RngString.Row).Value
                ActionName = RTrim(Left(ActionName, 8))
            Else
                ActionName = "Action not found"
            End If
            
            '-----------------------De standaard reden opzoeken
            If Not RngString Is Nothing Then
                DefaultReason = Sheets("DATA").Range("L" & RngString.Row).Value
            Else
                DefaultReason = "Action not found"
            End If
            
            Select Case Target.Value
                Case 3, 4, 5, 11 'de actie is: Toevoegen, -/-Rol, -/- cert of On Hold
                    ActionComment = InputBox(Prompt:="Give your reason for the following action:" & vbNewLine _
                        & "     " & ActionName, Title:="Give reason for" & ActionName, _
                        Default:=DefaultReason)

                    If ActionComment = vbNullString Then
                       Exit Sub

                    Else
                        If IsEmpty(Range("L" & Target.Row).Value) Then
                            Range("L" & Target.Row).Value = ActionName & "> " & ActionComment
                        Else
                            Range("L" & Target.Row).Value = ActionName & "> " & ActionComment & " | " & Range("L" & Target.Row).Value
                        End If
                    End If
            End Select
        End If
    End If
End If

End Sub

