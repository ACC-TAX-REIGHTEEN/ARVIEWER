Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("N2")) Is Nothing Then
        On Error GoTo SafeExit
        Application.EnableEvents = False
        Call UpdateDataAnalisis
        Application.EnableEvents = True
    End If

    If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        On Error GoTo SafeExit
        Application.EnableEvents = False
        Call UpdateDataAnalisisB
        Application.EnableEvents = True
    End If
    
    If Not Intersect(Target, Me.Range("Z2")) Is Nothing Then
        On Error GoTo SafeExit
        Application.EnableEvents = False
        Call UpdateDataAnalisisC
        Application.EnableEvents = True
    End If
    
    Exit Sub

SafeExit:
    Application.EnableEvents = True
End Sub
