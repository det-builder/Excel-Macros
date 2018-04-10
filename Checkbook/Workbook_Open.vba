Private Sub Workbook_Open()
Dim lastRow As Long
Dim scrlrw As Long
Dim activeCell As Range

    lastRow = wksRegister.Cells(500000, 2).End(xlUp).Row
    Set activeCell = wksRegister.Cells(lastRow + 1, 2)
    scrlrw = IIf(lastRow > 15, lastRow - 15, 1)
    
    Application.Goto reference:=activeCell
    ActiveWindow.ScrollRow = scrlrw

End Sub
