Attribute VB_Name = "SaveData"
Sub SaveOldData()

SubName = "'SaveOldData'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

CertBewerkbaar

Dim Sht As Worksheet
Dim i As Integer
Dim rng As Range
Dim Nm As String

'Ingevulde informatie bewaren
'Check on previouse saved data
    Sheets("Certificaten").Select
    Nm = Format(Range("A1").Value, "mm-dd-yyyy")

If Nm <> "" Then

For Each Sht In Worksheets
If Sht.Name = Nm Then

GoTo ClearSheet:
End If
Next Sht

GoTo NewSheet:

Else
BackgroundFunction.AutoCloseMessage Tekst:="There is no valid date in Range(A1)", Titel:="Task is canceld"
Exit Sub
End If

NewSheet:
'Make new sheet
    BackgroundFunction.AutoCloseMessage Tekst:="A new datasheet is made with the following date: " & Nm & "(mm-dd-yyyy)"
    Sheets.Add.Name = Nm
    
    Sheets("Certificaten").Select

    Range("A1:L1").Copy
    
    Sheets(Nm).Range("A1").PasteSpecial (xlPasteAll)
    
    GoTo CopyData:
Exit Sub

ClearSheet:

    ClearSheet (Nm)
    
    GoTo CopyData:

Exit Sub

CopyData:
'Copy existing data

    Sheets("Certificaten").Select

    Einde = Range("C1000").End(xlUp).Row

    Set rng = Range("A2", "G" & Einde)
    
    rng.Copy

'Paste existing data
    Sheets(Nm).Visible = xlSheetVisible
    Sheets(Nm).Select
    Range("A2").PasteSpecial xlPasteValues

'Second copy existing data
    Sheets("Certificaten").Select
    
    Set rng = Range("L2", "L" & Einde)
    
    rng.Copy
    
'Second paste existing data
    Sheets(Nm).Select
    Range("L2").PasteSpecial xlPasteValues

'Back to workable sheet
    Sheets("Certificaten").Select
    Range("A1").Select
    CertNietBewerkbaar
    
    Sheets(Nm).Visible = xlSheetVeryHidden

    HideAllSheets

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Sub SavePDF()

SubName = "'SavePDF'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

EndCert = Range("C1000").End(xlUp).Row

Range("A2", "L" & EndCert). _
        ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "J:\Certificaten\Certificaten Aflopend" & Format(Range("A1").Value, "mm-dd-yyyy") & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
Range("A2").Select

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub
