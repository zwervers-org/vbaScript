Attribute VB_Name = "NotAvailable"
Sub NotAvailable()

SubName = "'NotAvailable'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

'gegevens opslaan in tijdelijk blad
'Sheets.Add
'With ActiveSheet
'    .Name = "NotAvailableData"
'    .Range("A1").Value = "Code"      'Code
'    .Range("B1").Value = "Relatie"   'relatie
'    .Range("C1").Value = "Certificaat" 'certificaat
'    .Range("D1").Value = "EindDate"  'eind datum
'    .Range("E1").Value = "ContactAlias" 'contact persoon
'    .Range("F1").Value = "BeschikDate"    'beschikbaarheidsdatum
'    .Range("G1").Value = "ExtraInfo"  'extra info
'End With

'gegevens van certifaten downloaden met vervaldatum na nu-7dagen -/- status verwerkt
Sheets("NotAvailable").Visible = xlSheetVisible
Sheets("Certificaten").Visible = xlSheetVisible

Sheets("NotAvailable").Range("q2", "w" & Sheets("NotAvailable").Range("Q10000").End(xlUp).Row).Clear

Set rng = Sheets("Certificaten").Range("A2", "L" & Sheets("Certificaten").Range("C1000").End(xlUp).Row)

For Each rw In rng.Rows
    If InArray("NotAv", Sheets("Certificaten").Range("A" & rw.Row).Value) = True Then
        If CDate(Sheets("Certificaten").Range("I" & rw.Row).Value) < Format(Now() - 7, "d-mm-yyyy") Then
            Sheets("Certificaten").Range("C" & rw.Row & ":D" & rw.Row & ",G" & rw.Row & ",I" & rw.Row & ",L" & rw.Row).Copy
            
            RowNew = Sheets("NotAvailable").Range("Q10000").End(xlUp).Row + 1
            Sheets("NotAvailable").Range("Q" & RowNew).PasteSpecial xlPasteValues
        End If
    End If
Next rw

NotAv.Show
'dropdown list box voor naam contact persoon
'-->in form fill list SelectContact

'wanneer nog niet beschikbaar toevoegen
    'vragen naar naam
    'vragen naar email
    'vragen naar telefoonnr
'voor elke relatie dezelfde contactpersoon?
    'anders weergeven welke tijdelijke certifaten er aangemaakt moeten worden [selectlist]
        'laten selecteren welke dit contactpersoon kunnen hebben, rest nog een keer vragen.
    

'vragen naar verwachte beschikbaarheid document (datum)[met gegevens van het document/relatie]
'vragen naar extra informatie
'laten selecten waarvoor dit geld of vink met alles appart

Sheets("NotAvailable").Visible = xlSheetVeryHidden
Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Sub
