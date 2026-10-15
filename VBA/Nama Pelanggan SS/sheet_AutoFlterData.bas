Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B1")) Is Nothing Then Exit Sub
    
    Dim barisAkhir As Long
    Dim i As Long
    Dim kataKunci As String
    Dim teksB As String, teksC As String
    Dim rangeSembunyi As Range
    
    kataKunci = LCase(Me.Range("B1").Value)
    barisAkhir = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Me.Rows("5:" & barisAkhir).Hidden = False
    
    If kataKunci <> "" Then
        For i = 3 To barisAkhir
            teksB = LCase(Me.Cells(i, 2).Value)
            teksC = LCase(Me.Cells(i, 3).Value)
            
            If InStr(1, teksB, kataKunci) = 0 And InStr(1, teksC, kataKunci) = 0 Then
                If rangeSembunyi Is Nothing Then
                    Set rangeSembunyi = Me.Rows(i)
                Else
                    Set rangeSembunyi = Union(rangeSembunyi, Me.Rows(i))
                End If
            End If
        Next i
        
        If Not rangeSembunyi Is Nothing Then
            rangeSembunyi.EntireRow.Hidden = True
        End If
    End If
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
