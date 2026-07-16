Sub SalinDariComposed()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    
    Set wsSource = ThisWorkbook.Sheets("Composed")
    Set wsDest = ThisWorkbook.Sheets("Extracted")
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo Cleanup
    
    wsSource.Range("N4:X110").Copy
    wsDest.Range("B4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    
    wsDest.Select
    wsDest.Range("B4").Select

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
