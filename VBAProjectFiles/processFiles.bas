Attribute VB_Name = "processFiles"
Option Explicit
Dim curWb As Workbook
Dim curWs As Worksheet
Dim newWb As Workbook
Dim destWb As Workbook
Dim destWs As Worksheet
Dim newWs As Worksheet

Sub process()
    Dim proceed As Boolean
    Dim fileNames As String
    Dim strCopyRange As String
    Set curWb = ThisWorkbook
    Set curWs = curWb.Sheets("Main")
    If Not InStr(1, curWs.Range("c3").Value, Format(Date - 1, "yyyymmdd")) > 0 Or curWs.Range("c3").Value = "" Then
        proceed = True
        If fileNames <> "" Then
            fileNames = fileNames & "," & "IBFile"
        Else
            fileNames = "IBFile"
        End If
    End If
    If Not InStr(1, curWs.Range("c5").Value, Format(Date - 1, "yyyymmdd")) > 0 Or curWs.Range("c5").Value = "" Then
        proceed = True
        If fileNames <> "" Then
            fileNames = fileNames & "," & "Cash File"
        Else
            fileNames = "Cash File"
        End If
    End If
    If Not InStr(1, curWs.Range("c7").Value, Format(Date - 1, "yyyymmdd")) > 0 Or curWs.Range("c7").Value = "" Then
        proceed = True
        If fileNames <> "" Then
            fileNames = fileNames & "," & "Position File"
        Else
            fileNames = "Position File"
        End If
    End If
    If Not InStr(1, curWs.Range("c9").Value, Format(Date - 1, "yyyymmdd")) > 0 Or curWs.Range("c9").Value = "" Then
        proceed = True
        If fileNames <> "" Then
            fileNames = fileNames & "," & "TXS File"
        Else
            fileNames = "TXS File"
        End If
    End If
    If proceed = True Then
        If MsgBox("Recent files not found in " & fileNames & Chr(10) & "Proceed?", vbCritical + vbYesNo, "Process PNL") = vbNo Then
            Exit Sub
        End If
    End If
    If openPNL_file = True Then
        Application.ScreenUpdating = False
        processFile curWs.Range("c5").Value, "ABN Input", "A:M", "A1"
        processFile curWs.Range("c7").Value, "ABN Input", "I:V", "S1"
        processFile curWs.Range("c9").Value, "ABN Input", "N:AB", "AL1"
        
        destWs.Range("BS9:DR9").Copy
        Set newWs = destWb.Sheets("ABN Merge")
        strCopyRange = "B" & newWs.Range("a65536").End(xlUp).Row + 1
        If newWs.Range(strCopyRange).Offset(-1, -1).Value = Date - 1 Then
            If MsgBox("Data already exists, Yes to Replace or No to New Entry?", vbCritical + vbYesNo, "Data Already Exists") = vbNo Then
                newWs.Range(strCopyRange).PasteSpecial xlPasteValues
                newWs.Range(strCopyRange).Offset(0, -1).Value = Date - 1
            Else
                strCopyRange = Range(strCopyRange).Offset(-1, 0).Address
                newWs.Range(strCopyRange).PasteSpecial xlPasteValues
                newWs.Range(strCopyRange).Offset(0, -1).Value = Date - 1
            End If
        Else
            strCopyRange = Range(strCopyRange).Offset(-1, 0).Address
            newWs.Range(strCopyRange).PasteSpecial xlPasteValues
            newWs.Range(strCopyRange).Offset(0, -1).Value = Date - 1
        
        End If
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
        newWs.Activate
    End If
    MsgBox "Files processed successfully"
    
End Sub
Function openPNL_file() As Boolean
    Application.DisplayAlerts = False
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select the Daily P & L File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files Only", "*.xls*"
    If .Show <> -1 Then
        MsgBox "No Files selected, Exiting Program"
        openPNL_file = False
        Exit Function
    End If
    Set destWb = Workbooks.Open(.SelectedItems(1))
    End With
    openPNL_file = True
    Application.DisplayAlerts = True
    
End Function
Function processFile(fileName As String, destShtname As String, copyRange As String, pasteRange As String)
    Set destWs = destWb.Sheets(destShtname)
    Set newWb = Application.Workbooks.Open(fileName)
    Set newWs = newWb.Sheets(1)
    newWs.Range(copyRange).Copy
    destWs.Range(pasteRange).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    newWb.Saved = True
    newWb.Close False
End Function
Function processIBFile() As Boolean
    Set newWb = Application.Workbooks.Open(curWs.Range("c3").Value)
    
End Function

