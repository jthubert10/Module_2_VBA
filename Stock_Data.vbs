Attribute VB_Name = "Module1"
Sub Stock()

    'Apply to each worksheet
    For Each ws In Worksheets

    'Insert headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
      
      'Set variables for ticker symbol and total of that ticker symbol
        Dim Ticker As String
        Dim Total As Double
      
      'Set variables for the greatest table values
        Dim greatestDecrease As Double
        Dim greatestIncrease As Double
        Dim greatestTotal As Double
        Dim greatestDecrease_Ticker As Variant
        Dim greatestIncrease_Ticker As Variant
        Dim greatestTotal_Ticker As Variant

        'Set variables for range of greatest table
        Dim tickerTable As Range
        Set tickerTable = ws.Range("I2:I2836")
        Dim percentTable As Range
        Set percentTable = ws.Range("K2:K2836")
        Dim totalTable As Range
        Set totalTable = ws.Range("L2:L2836")

        'Set lr as last row in table
        Dim lr As Long
        lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set total volume to zero
        Total = 0

        'Set counter for use later
        Dim counter As Integer
        counter = 2

        'Identify and set open and close price variables/value
        Dim openPrice As Double
        Dim closePrice As Double
     
        openPrice = ws.Cells(2, 3).Value

        'Start for loop
        For i = 2 To lr
            
            Total = Total + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
    
                ws.Range("I" & counter).Value = Ticker
    
                ws.Range("L" & counter).Value = Total
          
                closePrice = ws.Cells(i, 6).Value
    
                Dim yearlyChange As Double
                Dim percentChange As Double
         
                yearlyChange = closePrice - openPrice
          
            If openPrice <> 0 Then

                percentChange = yearlyChange / openPrice
          
            Else: percentChange = 0
          
          End If
          
          openPrice = ws.Cells(i + 1, 3).Value
    
          ws.Range("J" & counter).Value = yearlyChange
         
          ws.Range("K" & counter).Value = percentChange
          
          If yearlyChange < 0 Then
        
            ws.Range("J" & counter).Interior.ColorIndex = 3
        
            Else: ws.Range("J" & counter).Interior.ColorIndex = 4
          
          End If
    
          counter = counter + 1
          
          Total = 0
    
        End If
     
      Next i
      
    'Find minimum and maximum values
    greatestIncrease = Application.Max(percentTable)
    greatestDecrease = Application.Min(percentTable)
    greatestTotal = Application.Max(totalTable)
    
    'Find matching tickers with the minumum and maximum values just found
    greatestIncrease_Ticker = WorksheetFunction.Index(tickerTable, WorksheetFunction.Match(greatestIncrease, percentTable, 0))
    greatestDecrease_Ticker = WorksheetFunction.Index(tickerTable, WorksheetFunction.Match(greatestDecrease, percentTable, 0))
    greatestTotal_Ticker = WorksheetFunction.Index(tickerTable, WorksheetFunction.Match(greatestTotal, totalTable, 0))
    
    'Populate bonus table with min/max macro
    ws.Range("Q2").Value = greatestIncrease
    ws.Range("Q3").Value = greatestDecrease
    ws.Range("Q4").Value = greatestTotal
    
    'Populate bonus table with matching ticker information
    ws.Range("P2").Value = greatestIncrease_Ticker
    ws.Range("P3").Value = greatestDecrease_Ticker
    ws.Range("P4").Value = greatestTotal_Ticker
    
    'Format cells and cell width
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    percentTable.NumberFormat = "0.00%"
    ws.Range("A:Q").Columns.AutoFit

Next ws

End Sub




