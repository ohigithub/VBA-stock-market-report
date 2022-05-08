# VBA-challenge

Sub stock2():

'Set initial variable for holding the ticker name
Dim Ticker As String

Dim Ticker_Total As Double
Ticker_Total = 0

Range("i1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Track location for each ticker name in the new ticker row
Dim New_Ticker_Row As Integer
New_Ticker_Row = 2

Dim OpenPrice As Long
OpenPrice = 2

'Loop through all ticker names to output the symbol
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To Lastrow


    'See if we are still within the same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'Set ticker name
    Ticker = Cells(i, 1).Value

    ' Print the ticker name in the new Ticker Column
    Range("I" & New_Ticker_Row).Value = Ticker
    
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    Range("L" & New_Ticker_Row).Value = Ticker_Total
    
    Range("J" & New_Ticker_Row).Value = Cells(i, 6).Value - Cells(OpenPrice, 3).Value
    
    Range("K" & New_Ticker_Row).Value = (Cells(i, 6).Value - Cells(OpenPrice, 3).Value) / Cells(OpenPrice, 3).Value
    
    New_Ticker_Row = New_Ticker_Row + 1
    
    Ticker_Total = 0
    
    OpenPrice = i + 1
    
    Else
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    End If

Next i

'Notate yearly change from opening price to the closing price for each year

Dim YearlyChange As Integer

For i = 2 To Lastrow
    
    
Next i

' Reset ticker total
Ticker_Total = 0

'Set an initial variable for holding the total volume of the stock

Ticker_Total = Ticker_Total + Cells(i, 7).Value

End Sub
