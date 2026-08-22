Sub SimpanDataKeDriveE_CustomName()
    Dim wsOrigin As Worksheet
    Dim wbNew As Workbook
    Dim rngData As Range
    Dim lastRow As Long
    Dim fileName As String
    Dim rawName As String
    Dim savePath As String
    
    On Error Resume Next
    Set wsOrigin = ThisWorkbook.Sheets("Extracted")
    On Error GoTo 0
    
    If wsOrigin Is Nothing Then
        MsgBox "Sheet bernama 'Extracted' tidak ditemukan!", vbCritical
        Exit Sub
    End If
    
    rawName = wsOrigin.Range("J4").Text & ", " & _
              wsOrigin.Range("L4").Text & ", " & _
              wsOrigin.Range("K4").Text & ", " & _
              wsOrigin.Range("N1").Text
    
    fileName = rawName
    fileName = Replace(fileName, "/", "-")
    fileName = Replace(fileName, "\", "-")
    fileName = Replace(fileName, ":", "-")
    fileName = Replace(fileName, "*", "")
    fileName = Replace(fileName, "?", "")
    fileName = Replace(fileName, """", "")
    fileName = Replace(fileName, "<", "(")
    fileName = Replace(fileName, ">", ")")
    fileName = Replace(fileName, "|", "-")
    
    If Trim(fileName) = ",,," Or Trim(fileName) = "" Then
        MsgBox "Nama file kosong. Pastikan J4, L4, K4, atau N1 ada isinya.", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsOrigin.Cells(wsOrigin.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 1 Then
        MsgBox "Tidak ada data untuk disimpan.", vbExclamation
        Exit Sub
    End If
    
    savePath = Range("L1").Value & fileName & ".xlsx"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set rngData = wsOrigin.Range("B3:O82")
    
    Set wbNew = Workbooks.Add
    rngData.Copy
    
    With wbNew.Sheets(1).Range("A1")
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With
    
    On Error Resume Next
    wbNew.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    
    If Err.Number <> 0 Then
        MsgBox "Gagal menyimpan file. Kemungkinan:" & vbCrLf & _
               "1. Drive E tidak ada." & vbCrLf & _
               "2. File sedang terbuka." & vbCrLf & _
               "3. Masih ada karakter aneh di nama file.", vbCritical
        wbNew.Close SaveChanges:=False
    Else
        wbNew.Close SaveChanges:=False
        MsgBox "File berhasil disimpan sebagai: " & vbCrLf & fileName & ".xlsx", vbInformation
    End If
    On Error GoTo 0
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
