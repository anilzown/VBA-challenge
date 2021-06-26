Attribute VB_Name = "Module1"
Sub Iterate()

'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox (LastRow)
'LastRow = 1000


Dim tickerCount  As Integer
Dim ticker As String
Dim openPrice As Double

tickerCount = 1
ticker = Cells(2, 1)
openPrice = Cells(2, 3)


Dim yearlyChange As Double
Dim percentChange As Double


Dim closingPrice As Double
Dim totalStockVolume As Double
totalStockVolume = 0

For i = 2 To LastRow

        If Cells(i, 1).Value = ticker Then
       
        closingPrice = Cells(i, 6)
        totalStockVolume = totalStockVolume + CDbl(Cells(i, 7).Value)
        
    Else
        ' Store the last ticker values in Grid
        'The ticker symbol.
        Cells(tickerCount + 1, 9).Value = ticker
        ticker = ""
        ' closing price =45.46   open price = 41.81
        'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        yearlyChange = closingPrice - openPrice
        Cells(tickerCount + 1, 10).Value = yearlyChange
    
                
        'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        Cells(tickerCount + 1, 15).Value = openPrice
        Cells(tickerCount + 1, 16).Value = yearlyChange
        
        Cells(tickerCount + 1, 17).Value = (yearlyChange / openPrice) * 100
        percentChange = (yearlyChange / openPrice) * 100
        Cells(tickerCount + 1, 11).Value = percentChange
        
        yearlyChange = 0
        percentChange = 0
        
        'The total stock volume of the stock.
        
        Cells(tickerCount + 1, 12).Value = totalStockVolume
        totalStockVolume = 0
                

        ' Ticker Inital Values Capture
        'totalStockVolume = 0
        tickerCount = tickerCount + 1
        ticker = Cells(i, 1).Value
        openPrice = Cells(i, 3).Value
        closingPrice = Cells(i, 6).Value
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
    End If
    Cells(i, 19).Value = i
    
Next
MsgBox (" Iam done")










End Sub



