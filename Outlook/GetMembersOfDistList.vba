Option Explicit

Public Sub GetMembers()
On Error GoTo catch_error

Const olFolderContacts = 10

Dim outlookApp As Outlook.Application
Dim outlookNamespace As Outlook.Namespace
Dim outlookAddrList As Outlook.addressList
Dim outlookEntry2 As Outlook.addressEntry
Dim outlookMember2 As Outlook.addressEntry
Dim intCount As Integer
Dim i As Integer
Dim rowToWriteTo As Integer
    
    ' Setup tasks.
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    rowToWriteTo = 2
    wksMembers.Cells.Select
    Selection.ClearContents
    Selection.Delete
    wksMembers.Range("A1").Select
    
    Debug.Print "Start " & CStr(Now)
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    Set outlookAddrList = outlookNamespace.AddressLists("Global Address List")
    Set outlookEntry2 = outlookAddrList.addressEntries("the_distribution_list")
    
    For i = 1 To outlookEntry2.Members.Count - 1
    'For i = 1 To 100
        Set outlookMember2 = outlookEntry2.Members.Item(i)
        wksMembers.Cells(rowToWriteTo, 1).Value = outlookMember2.Name
        rowToWriteTo = rowToWriteTo + 1
    Next
    
catch_error:

    On Error Resume Next
    Set outlookMember2 = Nothing
    Set outlookEntry2 = Nothing
    Set outlookAddrList = Nothing
    Set outlookNamespace = Nothing
    Set outlookApp = Nothing

    wksMembers.Range("A1").Value = "Display Name"
    wksMembers.Columns("A:A").EntireColumn.AutoFit
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    
    Debug.Print "Done " & CStr(Now)
       
    MsgBox "Done"
    
End Sub


