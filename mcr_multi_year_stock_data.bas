Attribute VB_Name = "Module1"
Sub Iterate()

'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

'Varibles to store the data column nubers
Dim TickerCellCol As Integer
Dim OpenPriceCellCol As Integer
Dim HighPriceCellCol As Integer
Dim LowPriceCellCol As Integer
Dim ClosePriceCellCol As Integer
Dim VolumeCellCol As Integer


'Variables to store the report column numbers
Dim TickerReportCellCol As Integer
Dim PerChangeReportCellCol As Integer
Dim YearlyChangeReportCellCol As Integer
Dim TotalStockVolumeReportCellCol As Integer

Dim LastRow As Long
Dim tickerCount  As Integer
Dim ticker As String
Dim openPrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim closingPrice As Double
Dim totalStockVolume As Double
    


TickerCellCol = 1
OpenPriceCellCol = 3
HighPriceCellCol = 4
LowPriceCellCol = 5
ClosePriceCellCol = 6
VolumeCellCol = 7

TickerReportCellCol = 9
YearlyChangeReportCellCol = 10
PerChangeReportCellCol = 11
TotalStockVolumeReportCellCol = 12
 
 For Each ws In Worksheets
 
 ' Make the worksheet active.
    ws.Activate
    
    ' last row calculation
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
      
    'MsgBox (LastRow)
    'LastRow = 1000

    'Step1 read the ticker symbol and initialize the variables with first row value.
    
    tickerCount = 1
    
    ticker = Cells(2, TickerCellCol)
    openPrice = Cells(2, OpenPriceCellCol)
 
    totalStockVolume = 0
        
    Cells(1, TickerReportCellCol).Value = "Ticker"
    Cells(1, YearlyChangeReportCellCol).Value = "Yearly Change"
    Cells(1, PerChangeReportCellCol).Value = "Percentage Change"
    Cells(1, TotalStockVolumeReportCellCol).Value = "Total Stock Volume"
    
    For i = 2 To LastRow
    
            If Cells(i, TickerCellCol).Value = ticker Then
           
            closingPrice = Cells(i, ClosePriceCellCol)
            totalStockVolume = totalStockVolume + CDbl(Cells(i, VolumeCellCol).Value)
            
        Else
            ' Store the last ticker values in Grid
            'The ticker symbol.
            Cells(tickerCount + 1, TickerReportCellCol).Value = ticker
            ticker = ""
            
            'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
            yearlyChange = closingPrice - openPrice
            Cells(tickerCount + 1, YearlyChangeReportCellCol).Value = yearlyChange
            
    
            'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            If openPrice > 0 Then
                percentChange = (yearlyChange / openPrice) '* 100
            Else
                percentChange = yearlyChange
            End If
            
            Cells(tickerCount + 1, PerChangeReportCellCol).Value = Format(percentChange, "Percent")
           
              If percentChange > 0 Then
                Cells(tickerCount + 1, PerChangeReportCellCol).Interior.ColorIndex = 4
            Else
                ' If yearly change value is less than 0, shade cell red.
                Cells(tickerCount + 1, PerChangeReportCellCol).Interior.ColorIndex = 3
           End If
            
            yearlyChange = 0
            percentChange = 0
            
            'The total stock volume of the stock.
            Cells(tickerCount + 1, TotalStockVolumeReportCellCol).Value = totalStockVolume
            
            totalStockVolume = 0
                    
    
            ' Ticker Inital Values Capture
            tickerCount = tickerCount + 1
            
            ticker = Cells(i, TickerCellCol).Value
            openPrice = Cells(i, OpenPriceCellCol).Value
            closingPrice = Cells(i, ClosePriceCellCol).Value
            
            totalStockVolume = totalStockVolume + Cells(i, VolumeCellCol).Value
            
        End If
        
        
    Next
    
    Dim GreatestPerIncreaseTickerName As String
    Dim GreatestPerIncreaseTickerValue As Double

    Dim GreatestPerDecreaseTickerName As String
    Dim GreatestPerDecreaseTickerValue As Double
    
    Dim GreatestVolumeTickerName As String
    Dim GreatestVolumeTickerValue As Double
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Get the last row
    LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
       
    'TickerReportCellCol = 9
    'YearlyChangeReportCellCol = 10
    'PerChangeReportCellCol = 11
    'TotalStockVolumeReportCellCol = 12
 
    ' Initialize variables and set values of variables initially to the first row in the list.
    GreatestPerIncreaseTickerName = Cells(2, TickerReportCellCol).Value
    GreatestPerIncreaseTickerValue = Cells(2, PerChangeReportCellCol).Value
    
    GreatestPerDecreaseTickerName = Cells(2, TickerReportCellCol).Value
    GreatestPerDecreaseTickerValue = Cells(2, PerChangeReportCellCol).Value
    
    GreatestVolumeTickerName = Cells(2, TickerReportCellCol).Value
    GreatestVolumeTickerValue = Cells(2, TotalStockVolumeReportCellCol).Value
    
    
    
    ' skipping the header row, loop through the list of tickers.
    For i = 2 To LastRow
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, PerChangeReportCellCol).Value > GreatestPerIncreaseTickerValue Then
            GreatestPerIncreaseTickerName = Cells(i, TickerReportCellCol).Value
            GreatestPerIncreaseTickerValue = Cells(i, PerChangeReportCellCol).Value
            
        End If
        
        Cells(i, PerChangeReportCellCol).Value = Cells(i, PerChangeReportCellCol).Value
        
        ' Find the ticker with the greatest percent decrease.
        If Cells(i, PerChangeReportCellCol).Value < GreatestPerDecreaseTickerValue Then
            GreatestPerDecreaseTickerName = Cells(i, TickerReportCellCol).Value
            GreatestPerDecreaseTickerValue = Cells(i, PerChangeReportCellCol).Value
            
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, TotalStockVolumeReportCellCol).Value > GreatestVolumeTickerValue Then
            GreatestVolumeTickerName = Cells(i, TickerReportCellCol).Value
            GreatestVolumeTickerValue = Cells(i, TotalStockVolumeReportCellCol).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet.
    Range("P2").Value = Format(GreatestPerIncreaseTickerName, "Percent")
    Range("Q2").Value = Format(GreatestPerIncreaseTickerValue, "Percent")
    
    Range("P3").Value = Format(GreatestPerDecreaseTickerName, "Percent")
    Range("Q3").Value = Format(GreatestPerDecreaseTickerValue, "Percent")
    
    Range("P4").Value = GreatestVolumeTickerName
    Range("Q4").Value = GreatestVolumeTickerValue
    
    
    
    
    ' -- MsgBox (" Iam done")
    
Next ws
    
    
    





End Sub



