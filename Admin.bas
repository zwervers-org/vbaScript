Attribute VB_Name = "Admin"
Sub ShowAllSaveSheet()

For i = 1 To Sheets.count

        If InArray("Sheets", Sheets(i).Name) Then
    
    Else
        Sheets(i).Visible = xlSheetVisible
    
    End If
    
Next i

End Sub
Sub ShowAllSheet()

For i = 1 To Sheets.count

        Sheets(i).Visible = xlSheetVisible
    
Next i

End Sub
Sub ShowOneSheet(Sht As String)

ActSht = ActiveSheet.Name
If ActSht <> Sht Then

        Sheets(Sht).Visible = xlSheetVisible
        Sheets(Sht).Select
        Sheets(ActSht).Visible = xlSheetVeryHidden
        
End If

End Sub

Sub HideAllSaveSheets()

For i = 1 To Sheets.count

    If InArray("Sheets", Sheets(i).Name) Then
        Sheets(i).Visible = xlSheetVisible
    Else
        Sheets(i).Visible = xlSheetVeryHidden
        
    End If
    
Next i

End Sub

Sub HideAllSheets()

With ThisWorkbook
    If .Sheets("Certificaten").Visible <> xlSheetVisible Then
        .Sheets("Certificaten").Visible = xlSheetVisible
        .Sheets("Certificaten").Select
    End If
    
    For i = 1 To .Sheets.count
        If .Sheets(i).Visible <> xlSheetVeryHidden Then
            If .Sheets(i).Name <> "Certificaten" Then
                .Sheets(i).Visible = xlSheetVeryHidden
            End If
        End If
    Next i
End With

End Sub

Sub HideOneSheet(Sht As String)

If Sht = "Certificaten" Or ActSht = "Certificaten" Then
Sheets("Certificaten").Select
GoTo Einde:

ElseIf ActSht = Sht Or ActSht = "" Then
ActSht = "Certificaten"

End If

    Sheets(ActSht).Visible = xlSheetVisible
    Sheets(ActSht).Select
    Sheets(Sht).Visible = xlSheetVerryHidden

Einde:
End Sub

Public Sub ExportVisualBasicCode()

SubName = "'ExportVisualBasicCode'"
If View("Errr") = True Then On Error GoTo ErrorText:

application.ScreenUpdating = View("Updte")
application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComp
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As Object
    
    'dirStart = ActiveWorkbook.path
    dirStart = "H:\ICT\Portable\Portable\PortableApps\GitPortable\App\Git" 'starting directory
    directory = "\VisualBasicScript\CertificatenAflopend" 'new directories for the vba-scripts
    fname = "Certificaten Aflopend.xlsm" 'this filename
    'path = dirStart & directory
    
    Set fso = CreateObject("scripting.filesystemobject")
    count = 0
    skiped = 0
    
    If Not fso.FolderExists(dirStart & directory) Then
        'when directory does not exists, make path
        newDir = dirStart
        Folders = Split(directory, "\")
        For i = 0 To UBound(Folders)
            newDir = fso.BuildPath(newDir, Folders(i))
            If fso.FolderExists(newDir) Then
                Set objFolder = fso.GetFolder(newDir)
            Else
                Set objFolder = fso.CreateFolder(newDir)
                Error.DebugTekst "Create folder: " & newDir
            End If
        Next
    End If
    Set fso = Nothing
    'Check if the right workbook is active
    If ActiveWorkbook.Name <> fname Then Workbooks(fname).Activate
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        Bladcheck = InStr(VBComponent.Name, "Blad")
        If Bladcheck > 0 Then
            Error.DebugTekst ("Skiped: " & VBComponent.Name)
            skiped = skiped + 1
            GoTo Volgende
        End If
        
        If count = 0 Then directory = dirStart & directory
        path = directory & "\" & VBComponent.Name & extension
        VBComponent.Export (path)
        If Err.Number <> 0 Then
            If InArray("VBAExport", extension) Then
                BackgroundFunction.AutoCloseMessage _
                        Tekst:="Failed to export: " & VBComponent.Name & vbNewLine _
                            & " to " & path & vbNewLine _
                            & vbNewLine & "Errornr: " & Err.Number & vbNewLine _
                            & "Description:  " & Err.Description, _
                        Titel:="Failed to export", _
                        Interval:=5, _
                        VoetTekst:=" "
                        
            End If
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If
        
Volgende:
        On Error GoTo ErrorText
    Next
    
    application.StatusBar = "Successfully exported: " & CStr(count) & " files | Skiped: " & CStr(skiped) & " files"
    Error.DebugTekst "Successfully exported: " & CStr(count) & " files | Skiped: " & CStr(skiped) & " files", SubName
   
'--------End Function
Error.DebugTekst Tekst:="Start GitHub"
Dim x As Variant
Dim ProgPath As String

' Set the Path variable equal to the path of your program's installation
    ProgPath = "H:\ICT\Portable\Portable\PortableApps\GitPortable\GitPortable.exe"

    x = Shell(ProgPath, vbNormalFocus)

Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Sub
