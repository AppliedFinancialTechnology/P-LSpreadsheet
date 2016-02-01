Attribute VB_Name = "importFiles"
Dim strFileName As String
Sub Unzip1(Fname As Variant)
    Dim FSO As Object
    Dim oApp As Object
    
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

        DefPath = StrReverse(Fname)
        DefPath = StrReverse(Replace(DefPath, Left(DefPath, InStr(1, DefPath, "\") - 1), ""))
        
        
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        FileNameFolder = DefPath
        Set oApp = CreateObject("Shell.Application")

        oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).items
    
        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.deletefolder Environ("Temp") & "\Temporary Directory*", True
    
End Sub

Public Function select_zipFiles(filestr As String) As String
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select the " & filestr & "File"
        .AllowMultiSelect = False
        .Filters.Clear
        If filestr = "IB" Then
            .Filters.Add "Excel Files Only", "*.xls; *.xlsx; *.xlsm", 1
            .Filters.Add "ZIP Files Only", "*.zip*; *.rar*", 2
        Else
            .Filters.Add "ZIP Files Only", "*.zip*; *.rar*"
        End If
        If .Show <> -1 Then
            MsgBox "No file selected", vbCritical + vbOKOnly
            select_zipFiles = ""
        Else
            
            select_zipFiles = .SelectedItems(1)
            
        End If
        
    End With
    If InStr(1, select_zipFiles, ".zip") > 0 Or InStr(1, select_zipFiles, ".rar") > 0 Then
        Unzip1 (select_zipFiles)
        select_zipFiles = Replace(Replace(Replace(select_zipFiles, ".zip", ""), ".rar", ""), "+", " ")
        If InStr(1, select_zipFiles, ".") = 0 Then
            select_zipFiles = select_zipFiles & ".xls"
            Exit Function
        End If
        If Len(select_zipFiles) - InStr(1, select_zipFiles, ".") > 3 Then select_zipFiles = Left(select_zipFiles, InStr(1, select_zipFiles, ".") + 3)
    End If
    
    
End Function

Public Sub ibImport()
    ThisWorkbook.Sheets(1).Range("ib_files").Value = select_zipFiles("IB")
End Sub
Public Sub cashImport()
    ThisWorkbook.Sheets(1).Range("cash_file").Value = select_zipFiles("Cash Summary")
End Sub
Public Sub positionImport()
    ThisWorkbook.Sheets(1).Range("position_file").Value = select_zipFiles("Position")
End Sub
Public Sub transImport()
    ThisWorkbook.Sheets(1).Range("trans_file").Value = select_zipFiles("Transaction")
End Sub
