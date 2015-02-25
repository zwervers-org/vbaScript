VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadDate 
   Caption         =   "Select date of data to load"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2655
   OleObjectBlob   =   "LoadDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub FindString_Click()

LoadDate.Hide

'    LoadOldData (FindString.Value)

FindString = FindString.Value

End Sub

Private Sub UserForm_Activate()

CertBewerkbaar

'Leeg beginnen
LoadDate.FindString.Clear

Einde = Range("AA1000").End(xlUp).Row 'einde zoeken
Range("AA1", "AB" & Einde + 5).Delete Shift:=xlUp

'Oude informatie ophalen

    For i = 1 To Sheets.count

        If InArray("Sheets", Sheets(i).Name) Then
    Else

        Cells(i, "AA").Value = Sheets(i).Name
        End If
    Next i

Einde = Range("AA1000").End(xlUp).Row 'einde resetten

'lege cellen verwijderen
For rw = Einde To 1 Step -1
If Cells(rw, "AA") = "" Then
Cells(rw, "AA").Delete Shift:=xlUp
End If
Next rw

'Sorteren op datum
    Einde = Range("AA1000").End(xlUp).Row 'einde resetten
    
    Range("AB1").FormulaR1C1 = _
        "=DATE(RIGHT(RC[-1],4),LEFT(RC[-1],2),MID(RC[-1],4,2))" 'datum juist zetten

    Range("AB1").AutoFill Destination:=Range("AB1:AB" & Einde), Type:=xlFillDefault

    Columns("AA:AB").Select
    ActiveWorkbook.Worksheets("Certificaten").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Certificaten").Sort.SortFields.Add Key:=Range( _
        "AB1:AB" & Einde + 100), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Certificaten").Sort
        .SetRange Range("AA1:AB" & Einde + 100)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

CheckListType = Range("AC1").Value

If CheckListType <> "Large" Then 'alleen als er een kleine lijst gevraagd wordt
'Lijst kleiner maken
    Einde = Range("AA1000").End(xlUp).Row 'einde zoeken
    
    Range("AA1", "AB1").Delete Shift:=xlUp 'huidige verwijderen
    
    Range("AA10", "AB" & Einde).Delete Shift:=xlUp 'laatste tien data overhouden
End If

'Zet alles in de selectielijst
    Einde = Range("AA1000").End(xlUp).Row 'einde resetten

    With FindString
    For i = 1 To Sheets("Certificaten").Range("AA1", "AA" & Einde).count
    
    .AddItem Range("AA" & i).Value
    
    Next i
    End With

'Leeg maken
Sheets("Certificaten").Range("AA1", "AC" & Einde).ClearContents

CertNietBewerkbaar

End Sub

