Sub SalinFormatKhususClipboard()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim targetSheets As Variant
    Dim lastRow As Long, i As Long, r As Long
    Dim noFaktur As String, nilaiFaktur As String, tglStr As String
    Dim tglFaktur As Date, selisihHari As Long
    Dim isFound As Boolean
    Dim tsheet As Variant
    Dim checkCell As String
    Dim lastRowTarget As Long
    
    Dim teksAkhir As String
    teksAkhir = ""
    
    Set wsSource = ActiveSheet
    
    targetSheets = Array("SOLO", "YOGYA", "SEMARANG")
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "Tidak ada data mulai dari baris 4!", vbExclamation
        Exit Sub
    End If
    
    For i = 4 To lastRow
        noFaktur = Trim(CStr(wsSource.Cells(i, "B").Value))
        
        If noFaktur <> "" Then
            
            nilaiFaktur = Format(wsSource.Cells(i, "G").Value, "#,##0")
            
            tglStr = Trim(wsSource.Cells(i, "C").Text)
            selisihHari = 0
            
            tglFaktur = ParseIndonesianDate(tglStr)
            
            If tglFaktur <> #1/1/1900# And tglFaktur <> #12:00:00 AM# Then
                selisihHari = DateDiff("d", tglFaktur, Date)
            End If
            
            isFound = False
            For Each tsheet In targetSheets
                On Error Resume Next
                Set wsTarget = ThisWorkbook.Sheets(tsheet)
                On Error GoTo 0
                
                If Not wsTarget Is Nothing Then
                    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row
                    
                    For r = 1 To lastRowTarget
                        checkCell = CStr(wsTarget.Cells(r, "D").Value)
                        
                        If InStr(1, checkCell, noFaktur) > 0 Then
                            isFound = True
                            Exit For
                        End If
                    Next r
                End If
                
                If isFound Then Exit For
            Next tsheet
            
            Dim ketHari As String
            If isFound Then
                ketHari = selisihHari & " HR (OWING)"
            Else
                ketHari = selisihHari & " HR"
            End If
            
            teksAkhir = teksAkhir & noFaktur & vbTab & " " & nilaiFaktur & " " & vbTab & " " & ketHari & vbCrLf
            
        End If
    Next i
    
    If teksAkhir <> "" Then
        Dim MyData As Object
        On Error Resume Next
        Set MyData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        On Error GoTo 0
        
        If Not MyData Is Nothing Then
            MyData.SetText teksAkhir
            MyData.PutInClipboard
            MsgBox "Data berhasil diformat dan disalin ke Clipboard!" & vbCrLf & _
                   "Silakan gunakan Ctrl + V untuk paste di mana saja.", vbInformation, "Sukses Salin"
        Else
            MsgBox "Gagal mengakses Clipboard sistem.", vbCritical
        End If
    Else
        MsgBox "Tidak ada data yang berhasil diproses.", vbExclamation
    End If
End Sub

Function ParseIndonesianDate(ByVal txt As String) As Date
    On Error GoTo InCaseError
    
    If IsDate(txt) Then
        ParseIndonesianDate = CDate(txt)
        Exit Function
    End If
    
    Dim komponen() As String
    komponen = Split(txt, " ")
    
    If UBound(komponen) = 2 Then
        Dim d As Integer, m As Integer, y As Integer
        d = CInt(komponen(0))
        y = CInt(komponen(2))
        
        Select Case LCase(Left(komponen(1), 3))
            Case "jan": m = 1
            Case "feb": m = 2
            Case "mar": m = 3
            Case "apr": m = 4
            Case "mei", "may": m = 5
            Case "jun": m = 6
            Case "jul": m = 7
            Case "agu", "aug": m = 8
            Case "sep": m = 9
            Case "okt", "oct": m = 10
            Case "nop", "nov": m = 11
            Case "des", "dec": m = 12
            Case Else: m = 0
        End Select
        
        If m > 0 Then
            ParseIndonesianDate = DateSerial(y, m, d)
            Exit Function
        End If
    End If

InCaseError:
    ParseIndonesianDate = #1/1/1900#
End Function
