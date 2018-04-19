Option Explicit

' Member variables
Private theFundID As String
Private theUserAgent As String
Private theURL As String
Private theAmt As Double

' Class constructor(s)
Private Sub Class_Initialize()
    theAmt = 0
    theFundID = ""
    theURL = "http://bigcharts.marketwatch.com/advchart/frames/frames.asp?symb="
    theUserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:59.0) Gecko/20100101 Firefox/59.0"
End Sub

' Class Methods.
Public Sub DoWork()
Dim xmlRequest As New WinHttp.WinHttpRequest
Dim theXML As String
Dim thePos1 As Integer
Dim stringSearch1 As String
Dim thePos2 As Integer
Dim stringSearch2 As String
Dim thePos3 As Integer
Dim theAmount As String

    xmlRequest.Option(WinHttpRequestOption_UserAgentString) = theUserAgent
    xmlRequest.Open "GET", theURL & theFundID, True
    xmlRequest.send
    xmlRequest.waitForResponse
    theXML = xmlRequest.responseText
    Set xmlRequest = Nothing
    
    stringSearch1 = "class=" & Chr(34) & "last" & Chr(34)
    thePos1 = InStr(1, theXML, stringSearch1)
    thePos2 = InStr(thePos1, theXML, "<div>")
    thePos3 = InStr(thePos2, theXML, "</div>")
    theAmount = Mid(theXML, thePos2 + 5, thePos3 - thePos2 - 5)

    If IsNumeric(theAmount) = True Then
        theAmt = CDbl(theAmount)
    End If
    
End Sub

' Properties
Property Get GivenPrice() As Double
    GivenPrice = theAmt
End Property

Property Let FundId(value As String)
    theFundID = value
End Property

