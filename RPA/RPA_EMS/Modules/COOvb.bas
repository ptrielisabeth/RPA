Sub DefaultPageFormat(sheetName)
    Sheets(sheetName).Select
    ActiveWindow.FreezePanes = False
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
End Sub