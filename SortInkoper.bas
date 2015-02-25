Attribute VB_Name = "SortInkoper"
Sub InkoperSorteren()

SubName = "'InkoperSorteren'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

1
CertBewerkbaar

10
Sheets("SortInk").Visible = xlSheetVisible

'werkblad schoonmaken
    Sheets("SortInk").Select
    
    ActiveWindow.View = xlNormalView
    
    Eind = Range("B1000").End(xlUp).Row + 50

    If Eind > 55 Then
        Columns("A:I").AutoFilter
        Sheets("SortInk").Range("A1", "P" & Eind).Clear
        Range("A1", "P" & Eind).Borders.LineStyle = xlNone
    End If

20 'copieer data
    Sheets("Certificaten").Select

    EindCert = Range("C1000").End(xlUp).Row
    ActiveSheet.Range("A1", "P" & EindCert).AutoFilter
    Range("A1", "L" & EindCert).Copy

30 'plak data op ander werkblad
    Sheets("SortInk").Range("A1").PasteSpecial xlPasteValues
    Sheets("SortInk").Select
    
40 'Delete overbodige informatie
    Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
    Columns("D:E").Select
        Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
        Selection.NumberFormat = "d/m/yyyy"
    
50 'Sorteer op inkoper
EindInk = Range("B1000").End(xlUp).Row

    Columns("A:I").Select
        ActiveWorkbook.Worksheets("SortInk").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("SortInk").Sort.SortFields.Add Key:=Range( _
            "G1", "G" & EindInk), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
         xlSortNormal
    With ActiveWorkbook.Worksheets("SortInk").Sort
        .SetRange Range("A1", "I" & EindInk)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Column1 = Range("A1").CurrentRegion.Columns.count
    If Column1 > 26 Then
    LastColumn = Chr(Int((Column1 - 1) / 26) + 64) & Chr(((Column1 - 1) Mod 26) + 65)
Else
    LastColumn = Chr(Column1 + 64)
End If

60 'Opmaak
    Range("A1", LastColumn & "1").Font.Bold = True
    With Range("A1", LastColumn & "1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Range("A1", LastColumn & EindInk).Font.Size = 8
    With Range("A1", LastColumn & EindInk + 1).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Range("H1", "H" & EindInk).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
61 '"Geen actie" verwijderen
    Set rng = Range("H2", "H" & EindInk)
    
    On Error Resume Next 'voorkomen dat "#NB" inhoud een foutmelding geeft
    
    For Each rw In rng
        If Range(rw.Address).Value = "Geen actie" Then
            Range(rw.Address).Value = ""
        End If
    Next rw

    On Error GoTo ErrorText
    
62 'Opmaak
    Columns("A:I").EntireColumn.AutoFit
    Columns("C:C").ColumnWidth = 23
    Columns("D:D").ColumnWidth = 23
    Columns("I:I").ColumnWidth = 40

63 'AutoFilter aanzetten op SortInk
    Range("A1", LastColumn & EindInk).AutoFilter
    
    Range("A1", LastColumn & EindInk).AutoFilter Field:=8, Criteria1:=Array( _
        "-/- certificaat", "-/- rol", "Aanvragen", "Controle", "Email", "Geen actie", _
        "Internet", ""), Operator:=xlFilterValues
        
    Range("A1").Select

64 'AutoFilter aanzetten op Certificaten
    Sheets("Certificaten").Select
    Cells.AutoFilter

70 'Set workmodus
    BackgroundFunction.CertNietBewerkbaar
    Admin.ShowOneSheet ("SortInk")
    'Sheets("SortInk").Visible = xlSheetVisible
    'Sheets("SortInk").Select
    'Sheets("Certificaten").Visible = xlSheetVerryHidden

Exit Sub
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Sub
