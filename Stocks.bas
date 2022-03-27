Attribute VB_Name = "Module1"
Sub Stocks()
    For Each ws In Worksheets
    'Create labels for each new column in the first row of each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Next ws
    
End Sub
