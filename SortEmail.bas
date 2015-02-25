Attribute VB_Name = "SortEmail"
Sub SorterenEmail(Data As String)

SubName = "'SorterenEmail'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")

CertBewerkbaar

10    Sheets(Data).Visible = xlSheetVisible
    Sheets(Data).Select
    
    EindeData = Range("A1000").End(xlUp).Row
    
20    If EindeData > 1 Then
    
    Range("A2", "Z" & EindeData).ClearContents

    End If

    Sheets("Certificaten").Select


    Einde = Range("C1000").End(xlUp).Row
    
30    Sheets("Certificaten").Range("A1", "Q" & Einde).AutoFilter

40    Sheets("Certificaten").Range("A1", "Q" & Einde).AutoFilter Field:=11, Criteria1:=Data
    
    EindeSelect = Range("C1000").End(xlUp).Row
        
50    If EindeSelect > 1 Then
        Sheets("Certificaten").Range("C2", "Q" & EindeSelect).Copy
    
        Sheets(Data).Select
    
60    Sheets(Data).Range("A2").PasteSpecial xlPasteValues
    
        EindeData = Range("A1000").End(xlUp).Row
        
        Range("H2", "I" & EindeData).Delete Shift:=xlToLeft
        
        Range("C2", "D" & EindeData).Delete Shift:=xlToLeft
        
        Range("F2", "F" & EindeData).Cut
        Range("L2").Insert Shift:=xlToRight
        
        SelectieCount = Range("A2", "A" & EindeData).count
    
    
70    Sheets("Certificaten").Select
    
71    Sheets("Certificaten").Range("A1", "Q" & Einde).AutoFilter Field:=11
    
    Sheets("Certificaten").Range("A1").Select
72    CertNietBewerkbaar
    
    Sheets(Data).Visible = xlSheetVisible
    Sheets(Data).Select
    Sheets("Certificaten").Visible = xlSheetVerryHidden
    
    Range("D2", "E" & Einde).NumberFormat = "d/m/yyyy"
    
    Sheets(Data).Range("L1").Select
    
    BackgroundFunction.AutoCloseMessage Tekst:=CountUnique(Sheets(Data).Range("A:A")) - 1 & _
        " adressen klaargezet voor emailen"
    
80    Else
    
81    ActiveSheet.Range("A1", "Q" & Einde).AutoFilter Field:=11
    Range("D2", "E" & Einde).NumberFormat = "m/d/yyyy"
    
    Range("A1").Select
    
    BackgroundFunction.AutoCloseMessage Tekst:="No data to email"
        
    End If
    
    Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Function CountUnique(ByVal rng As Range) As Long
Dim St As String
    Set rng = Intersect(rng, rng.Parent.UsedRange)
    St = "'" & rng.Parent.Name & "'!" & rng.Address(False, False)
    CountUnique = Evaluate("SUM(IF(LEN(" & St & "),1/COUNTIF(" & St & "," & St & ")))")
End Function
