Attribute VB_Name = "Module1"
Sub Stocks()
    For Each ws In Worksheets
    
    'Set a variable for holding the stock name
    Dim Ticker As String
    
    'Set a variable for holding the yearly change, opening and closing prices
    Dim Yearly_Change As Double
    Dim year_open As Double
    Dim year_close As Double
    
    
    'Set a variable for holding the percent change
    Dim Percent_Change As Double
    Percent_Change = 0
    
    'Set a variable for holding the total stock volume
    Dim Total_Volume As Double
    
    'Set a variable to keep track of the row for each stock
    Dim Stock_Row As Integer
    Stock_Row = 2
    
    'Set a variable for finding the last row of each sheet
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create labels for each new column in the first row of each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Create a for loop to read all of the data in each worksheet
    For i = 2 To LastRow
        
        'Conditional to find the opening price for each stock
        If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
            
            'Set variable for the opening price for each stock
            year_open = ws.Cells(i, 3).Value
            
        End If
        
        'Conditional for determining when we have read all data for each stock
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Set Ticker variable
            Ticker = ws.Cells(i, 1).Value
            
            'Set Total Volume variable
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            'Print Total Volume in column L
            ws.Range("L" & Stock_Row).Value = Total_Volume
            
            'Set year close variable to be the closing price in the last row for each stock
            year_close = ws.Cells(i, 6).Value
            
            'Calculate the yearly change
            Yearly_Change = year_close - year_open
            
            'Calculate the percent change
            Percent_Change = Yearly_Change / year_open
            
            'Print the yearly change in column J
            ws.Range("J" & Stock_Row).Value = Yearly_Change
            
            'Print the percent change in column K with percent formatting
            ws.Range("K" & Stock_Row).Value = FormatPercent(Percent_Change)
            
            'Print the ticker name in column I
            ws.Range("I" & Stock_Row).Value = Ticker
            
            'Increment the Stock Row variable so that each new stock and its data is printed in the correct row
            Stock_Row = Stock_Row + 1
            
            'Reset the Total Volume variable so that each stock's volume can be calculated separately
            Total_Volume = 0
            
        'Conditional to add the daily volume of each stock
        Else
            
            'Set total volume variable to add current day's volume to the running total
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
        End If
                
    Next i
    
    'For loop to read data from the newly created Yearly Change column
    For i = 2 To LastRow
        
        'Conditional to determine if a cell warrants a green fill
        If ws.Cells(i, 10).Value > 0 Then
            
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        'Conditional to determine if a cell warrants a red fill
        ElseIf ws.Cells(i, 10).Value < 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
    Next ws
    
End Sub
