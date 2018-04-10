Option Explicit

Private transDate As String
Private transCheckNum As String
Private transPayee As String
Private transCategory As String
Private transPayment As Currency
Private transClear As String
Private transMemo As String
Private transDeposit As Currency
Private cellRowNum As Integer
Private headerProcessed As Boolean

Private Enum ColumnLayout
    Date = 2
    CheckNumber = 4
    Payee = 6
    Category = 8
    Payment = 10
    ClrStatus = 12
    Deposit = 14
    Balance = 16
    Memo = 18
End Enum

Public Sub Import()
    ClearWorksheet
    ImportData
End Sub

Private Sub ClearWorksheet()
Dim lastRow As Long
Dim firstCell As Range
Dim lastCell As Range
Dim theRange As Range
    
    lastRow = wksRegister.Cells(500000, 2).End(xlUp).Row
    
    If lastRow > 1 Then
        Set firstCell = wksRegister.Range("B2")
        Set lastCell = wksRegister.Cells(lastRow, 100)
        Set theRange = wksRegister.Range(firstCell.Address & ":" & lastCell.Address)
    
        theRange.Select
        Selection.ClearContents
    End If
    
End Sub

Private Sub ImportData()

Dim lineOfData As String
Dim lineNumber As Long

PrepareApp (True)

lineNumber = 0
cellRowNum = 1
headerProcessed = False

WipeVariables
Open "D:\My_Downloads\My Money.QIF" For Input As #1
Line Input #1, lineOfData

Do While Not EOF(1)
    lineNumber = lineNumber + 1
    If lineNumber Mod 100 = 0 Then DoEvents
    
    ProcessLine (lineOfData)
    Line Input #1, lineOfData

Loop
ProcessLine (lineOfData)

Close #1
PrepareApp (False)
MsgBox "Done"

End Sub

Private Sub WipeVariables()
    transDate = ""
    transCheckNum = ""
    transPayee = ""
    transCategory = ""
    transPayment = 0
    transClear = ""
    transMemo = ""
    transDeposit = 0
End Sub

Private Sub PrepareApp(start As Boolean)

If start = True Then
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
Else
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
End If

End Sub

Private Sub ProcessLine(lineOfData As String)
   
    If Left(lineOfData, 1) = "^" Then
    
        cellRowNum = cellRowNum + 1
        wksRegister.Cells(cellRowNum, ColumnLayout.Date) = transDate
        wksRegister.Cells(cellRowNum, ColumnLayout.CheckNumber) = transCheckNum
        wksRegister.Cells(cellRowNum, ColumnLayout.Payee) = transPayee
        wksRegister.Cells(cellRowNum, ColumnLayout.Category) = transCategory
        wksRegister.Cells(cellRowNum, ColumnLayout.ClrStatus) = transClear
        
        If transPayment > 0 Then
            wksRegister.Cells(cellRowNum, ColumnLayout.Payment) = transPayment
        End If
        
        If transDeposit > 0 Then
            wksRegister.Cells(cellRowNum, ColumnLayout.Deposit) = transDeposit
        End If
        
        If headerProcessed = False Then
            wksRegister.Cells(cellRowNum, ColumnLayout.Balance) = transDeposit
            headerProcessed = True
        Else
            wksRegister.Cells(cellRowNum, ColumnLayout.Balance).FormulaR1C1 = "=(R[-1]C-RC[-6])+RC[-2]"
        End If
        
        wksRegister.Cells(cellRowNum, ColumnLayout.Memo) = transMemo
        
        WipeVariables
    
    ElseIf Left(lineOfData, 1) = "D" Then
        transDate = ParseDate(lineOfData)
    ElseIf Left(lineOfData, 1) = "T" Then
        If Mid(lineOfData, 2, 1) = "-" Then
            transPayment = CCur(Mid(lineOfData, 3))
        Else
            transDeposit = CCur(Mid(lineOfData, 2))
        End If
    ElseIf Left(lineOfData, 2) = "CX" Then
        transClear = "R"
    ElseIf Left(lineOfData, 1) = "N" Then
        transCheckNum = Mid(lineOfData, 2)
    ElseIf Left(lineOfData, 1) = "P" Then
        transPayee = Mid(lineOfData, 2)
    ElseIf Left(lineOfData, 1) = "M" Then
        transMemo = Mid(lineOfData, 2)
    ElseIf Left(lineOfData, 1) = "L" Then
        transCategory = Mid(lineOfData, 2)
    End If

End Sub




