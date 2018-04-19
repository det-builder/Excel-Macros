Option Explicit

' Miscrosoft XML v6.0 must be enabled from the VBA editor
' requires a reference to "Microsoft WinHTTP Services, version 5.1"

Public Function GetPrice(tickerID As String) As Double
Dim thePrice As New GetStockPrice

    thePrice.FundId = tickerID
    thePrice.DoWork
    GetPrice = thePrice.GivenPrice
    Set thePrice = Nothing
End Function
