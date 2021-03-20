Attribute VB_Name = "Module1"
Sub Stock_Analysis():

'Defining variables throughout the workbooks
Dim Ticker As String
Dim NumberTickers As Integer
Dim OpenPrice, ClosePrice As Double
Dim YearlyChange, PercentChange As Double
Dim TotalVolume As Double
Dim EndRow As Long

'Loop through each worksheet
For Each ws In Worksheets
    'Find the last row in each worksheet
    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Stating the additional headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Define starting count for each value
    NumberTickers = 0
    Ticker = ""
    YearlyChange = 0
    OpenPrice = 0
    PercentChange = 0
    TotalVolume = 0
    
    'Loop through the all tickers Symbols
    For i = 2 To EndRow
        
        'Give value to the ticker string
        Ticker = ws.Cells(i, 1).Value
        
        'Location of opening prices
        If OpenPrice = 0 Then
            OpenPrice = ws.Cells(i, 3).Value
        End If
        
        'Sum the total stock value for any given ticker
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        'Determine the entire breadth of one ticker. Combine, enter in column I, and move on to next ticker
        If ws.Cells(i + 1, 1).Value <> Ticker Then
            NumberTickers = NumberTickers + 1
            ws.Cells(NumberTickers + 1, 9) = Ticker
            
            'Location of closing prices
            ClosePrice = ws.Cells(i, 6)
           
            'Calculate the end of year difference in value
            YearlyChange = ClosePrice - OpenPrice
            
            'Distribute the yearly change of ticker's stockprice in column J and move on to the next ticker
            ws.Cells(NumberTickers + 1, 10).Value = YearlyChange
            
            'Color cells green for yearly change values greater than 0
            If YearlyChange > 0 Then
                ws.Cells(NumberTickers + 1, 10).Interior.ColorIndex = 4
            'Color cells red for yearly chaange values less than 0
            ElseIf YearlyChange < 0 Then
                ws.Cells(NumberTickers + 1, 10).Interior.ColorIndex = 3
            'Color cells yellow for yearly change values equal to zero
            Else
                ws.Cells(NumberTickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            'Determine a ticker's end of year change in value as a percent
            If OpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = (YearlyChange / OpenPrice)
            End If
                         
             'Apply formatting to display percent changed as a percentage
             ws.Cells(NumberTickers + 1, 11).Value = Format(PercentChange, "Percent")
           
            'Price must reset to zero when we reach the next ticker
            OpenPrice = 0
            
            'Reflect the annual sum total stock value for any given ticker in column L across all worksheets
            ws.Cells(NumberTickers + 1, 12).Value = TotalVolume
            
            'Volume must reset to zero when we reach the next ticker
            TotalVolume = 0
        End If
        
    Next i
    
    'Define next set of variables
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentIncreaseTicker As String
    Dim GreatestPercentDecrease As Double
    Dim GreatestPercentDecreaseTicker As String
    Dim GreatestStockVolume As Double
    Dim GreatestStockVolumeTicker As String
    
    'Create space to showcase the tickers with the greatest % increase, greatest % decrease, and greatest total volume
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    'Find the last row
    EndRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Define where to pull values to begin calculating the relevant data requests (% increase, decrease, etc)
    GreatestPercentIncrease = ws.Cells(2, 11).Value
    GreatestPercentIncreaseTicker = ws.Cells(2, 9).Value
    GreatestPercentDecrease = ws.Cells(2, 11).Value
    GreatestPercentDecreaseTicker = ws.Cells(2, 9).Value
    GreatestStockVolume = ws.Cells(2, 12).Value
    GreatestStockVolumeTicker = ws.Cells(2, 9).Value
    'Loop through to the last row
    For i = 2 To EndRow
        
        'Find the greatest percentage increase across all tickers
        If ws.Cells(i, 11).Value > GreatestPercentIncrease Then
            GreatestPercentIncrease = ws.Cells(i, 11).Value
            GreatestPercentIncreaseTicker = ws.Cells(i, 9).Value
        End If
        
        'Find the greatest percentage decrease
        If ws.Cells(i, 11).Value < GreatestPercentDecrease Then
            GreatestPercentDecrease = ws.Cells(i, 11).Value
            GreatestPercentDecreaseTicker = ws.Cells(i, 9).Value
        End If
        
        'Find the greatest stock volume
        If ws.Cells(i, 12).Value > GreatestStockVolume Then
            GreatestStockVolume = ws.Cells(i, 12).Value
            GreatestStockVolumeTicker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    'Assign the results of our analysis a destination
    ws.Range("P2").Value = Format(GreatestPercentIncreaseTicker, "Percent")
    ws.Range("Q2").Value = Format(GreatestPercentIncrease, "Percent")
    ws.Range("P3").Value = Format(GreatestPercentDecreaseTicker, "Percent")
    ws.Range("Q3").Value = Format(GreatestPercentDecrease, "Percent")
    ws.Range("P4").Value = GreatestStockVolumeTicker
    ws.Range("Q4").Value = GreatestStockVolume
        
Next ws


End Sub


