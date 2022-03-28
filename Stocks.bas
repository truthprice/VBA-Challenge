Attribute VB_Name = "Module1"
Sub Stocks()
    For Each ws In Worksheets
    
    'Set a variable for holding the stock name
    Dim Ticker As String
    
    'Set a variable for holding the yearly change
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    'Set a variable for holding the percent change
    Dim Percent_Change As Double
    Percent_Change = Percent_Change * 100
    
    'Set a variable for holding the total stock volume
    Dim Total_Volume As Double
    Total_Volume = 0
    
    Dim Stock_Row As Integer
    Stock_Row = 2
    
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Create labels for each new column in the first row of each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           Ticker = ws.Cells(i, 1).Value
            Total_Volume = Total_Volume + Cells(i, 7).Value
            ws.Range("I" & Stock_Row).Value = Ticker
            ws.Range("L" & Stock_Row).Value = Total_Volume
            Stock_Row = Stock_Row + 1
            Total_Volume = 0
            
        Else
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    Next ws
    
End Sub
