Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub GetData()
On Error GoTo catch_error

Dim outlookApp As Outlook.Application
Dim outlookAddressLists As Outlook.AddressLists
Dim addressList As Outlook.addressList
Dim addressEntries As Outlook.addressEntries
Dim addressEntry As Outlook.addressEntry
Dim exchangeUser As Outlook.exchangeUser
Dim oPA As Outlook.PropertyAccessor
Dim rowToWriteTo As Integer
Dim inputRecords As Long
Dim comments As String
Dim commentString As String
Dim operatorId As String
Dim country As String
Dim role As String
    
    ' Setup tasks.
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    rowToWriteTo = 2
    inputRecords = 0
    wksData.Cells.Select
    Selection.ClearContents
    Selection.Delete
    wksData.Range("A1").Select
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookAddressLists = outlookApp.Session.AddressLists
    
    Debug.Print "Start: " & CStr(Now)
    
    For Each addressList In outlookAddressLists
        If addressList.AddressListType = olExchangeGlobalAddressList Then
            Set addressEntries = addressList.addressEntries
            
            For Each addressEntry In addressEntries
                comments = ""
                country = ""
                inputRecords = inputRecords + 1
                
                DoEvents
                
                If inputRecords Mod 100 = 0 Then Sleep (1650)
               ' If inputRecords Mod 500 = 0 Then Sleep (10000)
                
                If inputRecords Mod 1000 = 0 Then
                    DoEvents
                    Sleep (1500)
                    DoEvents
                End If
                
                If addressEntry.AddressEntryUserType = olExchangeUserAddressEntry Then
                    Set exchangeUser = addressEntry.GetExchangeUser
                    
                    Set oPA = exchangeUser.PropertyAccessor
                    
                    On Error Resume Next
                    commentString = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3004001E") ' Comments
                    
                    operatorId = GetOperatorId(oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3004001E")) ' Operator Id
                    role = GetRole(commentString)
                    country = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A26001E") ' Country
                    
                    If operatorId = "" Or IsNull(operatorId) = True Or country <> "USA" Then
                        ' Continue on.
                    Else
                        wksData.Cells(rowToWriteTo, 1) = operatorId ' Operator Id
                        wksData.Cells(rowToWriteTo, 2) = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3001001E") ' Display Name
                        wksData.Cells(rowToWriteTo, 3) = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A16001E") ' Group Name
                        wksData.Cells(rowToWriteTo, 4) = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A28001E") ' State
                        wksData.Cells(rowToWriteTo, 5) = role ' Role
                        wksData.Cells(rowToWriteTo, 6) = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A17001E") ' Title
                        rowToWriteTo = rowToWriteTo + 1
                        
                    End If
                    
                    On Error GoTo catch_error
                    
                End If

                If rowToWriteTo >= 100 Then
                    GoTo catch_error
                End If
                
            Next
            
        End If
    Next
    
catch_error:

    On Error Resume Next
    Set exchangeUser = Nothing
    Set addressEntry = Nothing
    Set addressEntries = Nothing
    Set addressList = Nothing
    Set outlookAddressLists = Nothing
    Set outlookApp = Nothing

    wksData.Range("A1").Value = "Operator Id"
    wksData.Range("B1").Value = "Display Name"
    wksData.Range("C1").Value = "Department Name"
    wksData.Range("D1").Value = "State"
    wksData.Range("E1").Value = "Role"
    wksData.Range("F1").Value = "Title"
    
    wksData.Columns("A:A").EntireColumn.AutoFit
    wksData.Columns("B:B").EntireColumn.AutoFit
    wksData.Columns("C:C").EntireColumn.AutoFit
    wksData.Columns("D:D").EntireColumn.AutoFit
    wksData.Columns("E:E").EntireColumn.AutoFit
    wksData.Columns("F:F").EntireColumn.AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    
    Debug.Print "Done:  " & CStr(Now)
    
    MsgBox "Done"
    
End Sub

Private Function GetOperatorId(commentString As String) As String

Dim startingPosition As Integer
Dim endingPosition As Integer
Dim fullString1 As String

    commentString = Replace(commentString, Chr$(13), "")
    commentString = Replace(commentString, Chr$(10), "")
    
    startingPosition = InStr(1, commentString, "Operator ID:")
    endingPosition = InStr(startingPosition, commentString, "Service Date")
    fullString1 = Mid$(commentString, startingPosition, endingPosition - startingPosition)
    
    GetOperatorId = Replace(fullString1, "Operator ID: ", "")

End Function

Private Function GetRole(commentString As String) As String

Dim startingPosition As Integer
Dim fullString1 As String

    commentString = Replace(commentString, Chr$(13), "")
    commentString = Replace(commentString, Chr$(10), "")
    
    startingPosition = InStr(1, commentString, "Role:")
    fullString1 = Mid$(commentString, startingPosition + 6)
    
    GetRole = fullString1

End Function


