Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsBayar As Worksheet
    Dim wsSource As Worksheet
    Dim rngNoNota As Range
    Dim rngNominalBayar As Range
    Dim cell As Range
    Dim noFaktur As String
    Dim nilaiFaktur As Double
    Dim totalBayar As Double
    Dim jmlData As Long
    Dim lastRowBayar As Long
    Dim matchRow As Variant
    
    If Intersect(Target, Me.Columns(2)) Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo Cleanup
    
    Set wsBayar = ThisWorkbook.Sheets("Pembayaran SS")
    Set wsSource = ThisWorkbook.Sheets("Source")
    
    lastRowBayar = wsBayar.Cells(wsBayar.Rows.Count, "F").End(xlUp).Row
    If lastRowBayar < 2 Then lastRowBayar = 2
    
    Set rngNoNota = wsBayar.Range("F2:F" & lastRowBayar)
    Set rngNominalBayar = wsBayar.Range("H2:H" & lastRowBayar)
    
    For Each cell In Intersect(Target, Me.Columns(2))
        If cell.Row > 1 Then
            
            If Not IsEmpty(cell.Value) And cell.Value <> "" Then
                noFaktur = cell.Value
                
                nilaiFaktur = cell.Offset(0, 5).Value
                
                totalBayar = Application.WorksheetFunction.SumIf(rngNoNota, noFaktur, rngNominalBayar)
                
                jmlData = Application.WorksheetFunction.CountIf(rngNoNota, noFaktur)
                
                If jmlData > 0 Then
                    If totalBayar = nilaiFaktur Then
                        With cell.Offset(0, 13)
                            .Value = "LUNAS (Total " & jmlData & "x bayar)"
                            .Font.Color = vbBlue
                            .Font.Bold = True
                        End With
                    ElseIf totalBayar < nilaiFaktur Then
                        With cell.Offset(0, 13)
                            .Value = "DICICIL: Baru Bayar " & Format(totalBayar, "#,##0") & " (Kurang: " & Format(nilaiFaktur - totalBayar, "#,##0") & ")"
                            .Font.Color = RGB(255, 102, 0)
                            .Font.Bold = False
                        End With
                    Else
                        With cell.Offset(0, 13)
                            .Value = "LEBIH BAYAR: Total " & Format(totalBayar, "#,##0")
                            .Font.Color = vbMagenta
                            .Font.Bold = False
                        End With
                    End If
                Else
                    With cell.Offset(0, 13)
                        .Value = "Belum Ada Data"
                        .Font.Color = vbBlack
                        .Font.Bold = False
                    End With
                End If
                
                matchRow = Application.Match(cell.Value, wsSource.Columns(2), 0)
                
                If IsError(matchRow) And IsNumeric(cell.Value) Then
                    matchRow = Application.Match(Val(cell.Value), wsSource.Columns(2), 0)
                End If
                
                If IsError(matchRow) Then
                    matchRow = Application.Match(Trim(CStr(cell.Value)), wsSource.Columns(2), 0)
                End If
                
                If Not IsError(matchRow) Then
                    Me.Cells(cell.Row, 14).Value = wsSource.Cells(matchRow, 13).Value
                Else
                    Me.Cells(cell.Row, 14).Value = "Tidak Ditemukan"
                End If
                
            Else
                Me.Cells(cell.Row, 14).ClearContents
                cell.Offset(0, 13).ClearContents
            End If
            
        End If
    Next cell

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
