Option Explicit

Public Sub Tests()

    If ParseDate("D1/11/99") <> "01/11/1999" Then MsgBox "Error1"
    If ParseDate("D2/ 1/98") <> "02/01/1998" Then MsgBox "Error2"
    If ParseDate("D10/29/99") <> "10/29/1999" Then MsgBox "Error3"
    If ParseDate("D1/ 7' 0") <> "01/07/2000" Then MsgBox "Error4"
    If ParseDate("D1/15' 0") <> "01/15/2000" Then MsgBox "Error5"
    If ParseDate("D12/11' 0") <> "12/11/2000" Then MsgBox "Error6"
    If ParseDate("D4/ 1'11") <> "04/01/2011" Then MsgBox "Error7"

End Sub

Public Function ParseDate(lineOfData As String) As String

Dim month As String
Dim day As String
Dim year As String
Dim newLineOfData
Dim firstSlashPos As Integer
Dim secondSlashPos As Integer
        
    firstSlashPos = InStr(1, lineOfData, "/")
    month = Right("00" + LTrim(RTrim(Mid(lineOfData, 2, firstSlashPos - 2))), 2)
    
    day = Right("00" + LTrim(RTrim(Mid(lineOfData, firstSlashPos + 1, 2))), 2)
    
    secondSlashPos = InStr(firstSlashPos + 1, lineOfData, "/")
    If secondSlashPos > 0 Then
        year = "19" + Mid(lineOfData, secondSlashPos + 1)
    Else
        year = "20" + Right("00" + LTrim(RTrim(Right(lineOfData, 2))), 2)
    End If

    ParseDate = month + "/" + day + "/" + year

End Function
