VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NotAv 
   Caption         =   "UserForm1"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13215
   OleObjectBlob   =   "NotAv.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "NotAv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Function Standards(ByVal Whith As String)

Select Case Whith
Case "NotAvFalse"
    Standards = 220

Case "NotAv1"
    Standards = 440

Case "NotAv2"
    Standards = 665

Case "Frame1Pos"
    Standards = 5
    
Case "Frame2Pos"
    Standards = 225
    
Case "Frame3Pos"
    Standards = 450

End Select

End Function

Private Sub SaveAddButton_Click()

Nieuw = Sheets("NotAvailable").Range("J1000").End(xlUp).Row + 1

Sheets("NotAvailable").Range("J" & Nieuw).Value = AliasAdd.Value
Sheets("NotAvailable").Range("K" & Nieuw).Value = NameAdd.Value
Sheets("NotAvailable").Range("L" & Nieuw).Value = EmailAdd.Value
Sheets("NotAvailable").Range("M" & Nieuw).Value = PhoneAdd.Value

AddContactButton.Value = False

Call AddContactButton_Change

NotAv.StatusNotAv.Caption = "Contact " & Sheets("NotAvailable").Range("J" & Nieuw).Value & " is added"

AliasAdd.Value = "Short name for this contact"
NameAdd.Value = "Full name for this contact (First and family name)"
EmailAdd.Value = ""
PhoneAdd.Value = "+31 (0) "

End Sub

Private Sub SaveCertButton_Click()

Dim arrItems() As String

CertCount = 0

'NotAv.Width = Standards("NotAvFalse")

'NotAv.SelectCertFrame.Visible = False

Set rng = Sheets("NotAvailable").Range("Q2", "W" & Sheets("NotAvailable").Range("Q10000").End(xlUp).Row)

ReDim arrItems(0 To 0)

For i = 0 To NotAv.SelectCertBox.ListCount - 1
    If NotAv.SelectCertBox.Selected(i) Then
        For Each rw In rng.Rows
            If NotAv.SelectCertBox.Column(0, i) = Sheets("NotAvailable").Range("Q" & rw.Row) Then
                If NotAv.SelectCertBox.Column(2, i) = Sheets("NotAvailable").Range("S" & rw.Row) Then
                    If NotAv.SelectCertBox.Column(3, i) = Sheets("NotAvailable").Range("T" & rw.Row) Then
                                                
                        ReDim Preserve arrItems(0 To CertCount)
                        
                        arrItems(CertCount) = rw.Row
                        
                        CertCount = CertCount + 1
                    End If
                End If
            End If
        Next rw
    End If
Next i

For j = LBound(arrItems) To UBound(arrItems)
    Sheets("NotAvailable").Range("V" & arrItems(j)).Value = NotAv.SelectContact.Value
    Sheets("NotAvailable").Range("W" & arrItems(j)).Value = NotAv.CertExpect.Value
    
    'check if there is no info
    If Sheets("NotAvailable").Range("U" & arrItems(j)).Value = "" Then
    Sheets("NotAvailable").Range("U" & arrItems(j)).Value = NotAv.AddInfo.Value
    
    Else
    MsgBox "Er is al info in het vakje"
    End If
    
Next j

NotAv.StatusNotAv.Caption = "Information added to " & CertCount & " certificates"

NotAv.SelectContact.Value = ""
NotAv.CertExpect.Value = Format(Now(), "d-m-yyyy")
NotAv.AddInfo.Value = ""

End Sub

Function SwitchMessage(Button As String)

Select Case Button
'wanneer er op de MoreMessage wordt gedrukt
Case "MoreMessage"
    If SwitchMoreMessage.Value = True Then
        If NotAv.AddContactFrame.Visible = True Then
            NotAv.AddContactFrame.Left = Standards("Frame3Pos")
            NotAv.Width = Standards("NotAv2")
        Else
            NotAv.Width = Standards("NotAv1")
        End If
            
        NotAv.SelectCertFrame.Visible = True
        NotAv.SaveMain.Visible = False
        NotAv.SwitchOneMessage.Value = False
    
    Else
        If NotAv.AddContactFrame.Visible = True Then
            NotAv.AddContactFrame.Left = Standards("Frame2Pos")
            NotAv.Width = Standards("NotAv1")
        Else
            NotAv.Width = Standards("NotAvFalse")
        End If
    
        NotAv.SelectCertFrame.Visible = False
        NotAv.SaveMain.Visible = True
        NotAv.SwitchOneMessage.Value = True
    End If
    
Case "OneMessage"
    If SwitchOneMessage.Value = True Then
        
        If NotAv.AddContactFrame.Visible = True Then
            NotAv.AddContactFrame.Left = Standards("Frame2Pos")
            NotAv.Width = Standards("NotAv1")
        Else
            NotAv.Width = Standards("NotAvFalse")
        End If
            
        NotAv.SelectCertFrame.Visible = False
        NotAv.SaveMain.Visible = True
        NotAv.SwitchMoreMessage.Value = False
    
    Else
        If NotAv.AddContactFrame.Visible = True Then
            NotAv.AddContactFrame.Left = Standards("Frame3Pos")
            NotAv.Width = Standards("NotAv2")
        Else
            NotAv.Width = Standards("NotAv1")
        End If
    
        NotAv.SelectCertFrame.Visible = True
        NotAv.SaveMain.Visible = False
        NotAv.SwitchMoreMessage.Value = True
    End If

Case Else
    MsgBox "Problem 'SwitchMessager': can't select case"

End Select

Einde:
End Function

Private Sub AddContactButton_Change()

If NotAv.AddContactButton = True Then
    If NotAv.SelectCertFrame.Visible = True Then
        NotAv.AddContactFrame.Left = Standards("Frame3Pos")
        NotAv.Width = Standards("NotAv2")
    Else
        NotAv.Width = Standards("NotAv1")
    End If
    
    NotAv.AddContactFrame.Visible = True
    NotAv.SelectContact.Enabled = False

Else
    If NotAv.SelectCertFrame.Visible = True Then
        NotAv.AddContactFrame.Left = Standards("Frame2Pos")
        NotAv.Width = Standards("NotAv1")
    Else
        NotAv.Width = Standards("NotAvFalse")
    End If

    Eind = Sheets("NotAvailable").Range("J1000").End(xlUp).Row
    
    If Eind > 1 Then
        Set rng = Sheets("NotAvailable").Range("J2", "M" & Eind)
        
        With NotAv.SelectContact
            .ColumnCount = rng.Columns.count
            .ColumnHeads = True
            .RowSource = rng.Worksheet.Name & "!" & rng.Address
        End With
    End If
    
    NotAv.SelectContact.Enabled = True
    NotAv.AddContactFrame.Visible = False
End If
End Sub

Private Sub SwitchMoreMessage_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

Call SwitchMessage("MoreMessage")

End Sub

Private Sub SwitchOneMessage_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

Call SwitchMessage("OneMessage")

End Sub

Private Sub UserForm_Activate()

NotAv.Width = Standards("NotAvFalse")
NotAv.SelectCertFrame.Visible = False
NotAv.AddContactFrame.Visible = False
NotAv.SwitchOneMessage.Value = True
NotAv.SwitchMoreMessage.Value = False
NotAv.AddContactButton.Value = False

NotAv.SelectCertFrame.Left = Standards("Frame2Pos")
NotAv.AddContactFrame.Left = Standards("Frame2Pos")
NotAv.MainFrame.Left = Standards("Frame1Pos")

'Contact details load
Eind = Sheets("NotAvailable").Range("J1000").End(xlUp).Row
    
    If Eind > 1 Then
        Set rng = Sheets("NotAvailable").Range("J2", "M" & Eind)
        
        With NotAv.SelectContact
            .ColumnCount = rng.Columns.count
            .ColumnHeads = True
            .RowSource = rng.Worksheet.Name & "!" & rng.Address
        End With
    End If

'Load certificates

NotAv.SelectCertBox.Clear

    
Eind = Range("Q1000").End(xlUp).Row
        
    If Eind > 1 Then
        Set rng = Sheets("NotAvailable").Range("Q2", "W" & Eind)
        
        With NotAv.SelectCertBox
            .ColumnCount = rng.Columns.count
            .ColumnHeads = True
            .RowSource = rng.Worksheet.Name & "!" & rng.Address
        End With
    Else
    NotAv.SelectCertBox.AddItem ("NO DATA")
    End If
    
End Sub



