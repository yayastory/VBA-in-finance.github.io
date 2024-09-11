"Making the output currency dynamic"
Function ForextoUSD(AmountInEUR, InputCurrency)
Application.Volatile True
RateRangeName = "USDper" & InputCurrency
ForexRate = ws_Assumptions.Range(RateRangeName).Value
AmountInUSD = AmountInEUR * ForexRate
ForextoUSD = AmountInUSD

End Function
