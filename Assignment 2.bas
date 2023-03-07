Attribute VB_Name = "Module1"
Sub AllStocksAnalysis()
For Each ws In Worksheets
    'Figure Last Row
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Other Starting Variables
    output_row = 2
    symbol_vol = 0
    open_value = 2
    
    'Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    
    'Conditional Formatting for Yearly Change
   
    
    'Begin Loop
    For input_row = 2 To LastRow
        If ws.Cells(input_row + 1, 1).Value <> ws.Cells(input_row, 1).Value Then
        'Print Ticker
        ws.Cells(output_row, 9) = ws.Cells(input_row, 1).Value
        'Final Volume Calculation
        symbol_vol = symbol_vol + ws.Cells(input_row, 7).Value
        'Print Volume
        ws.Cells(output_row, 12).Value = symbol_vol
        'Print Change
        ws.Cells(output_row, 10).Value = ws.Cells(input_row, 6).Value - ws.Cells(open_value, 3)
        'Print Percent Change
        ws.Cells(output_row, 11).Value = ws.Cells(output_row, 10).Value / ws.Cells(open_value, 3)
        'Prepare for Next Ticker & Clean up
        output_row = output_row + 1
        symbol_vol = 0
        open_value = input_row + 1
        
        Else
        'Table Volume
            symbol_vol = symbol_vol + ws.Cells(input_row, 7).Value
        End If
    Next input_row
    
    'Stock Highlights
    
    'New Last Row
        LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    'Entries
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Find the Percentages
    'Set the Variables
        Dim Max_Percent As Double
        Dim Min_Percent As Double
        Max_Percent = 0
        Min_Percent = 0
        Max_Volume = 0
        Max_Stock = 0
        Min_Stock = 0
        Vol_stock = 0
    'Loop
    For i = 2 To LastRow2
    'Figure Max/Min Percent & Ticker Row
        If ws.Cells(i, 11).Value > Max_Percent Then
            Max_Percent = ws.Cells(i, 11).Value
            Max_Stock = i
        ElseIf ws.Cells(i, 11).Value < Min_Percent Then
            Min_Percent = ws.Cells(i, 11).Value
            Min_Stock = i
        End If
    'Figure Max Volume
        If ws.Cells(i, 12).Value > Max_Volume Then
            Max_Volume = ws.Cells(i, 12).Value
            Vol_stock = i
        End If
    'Formatting
        If ws.Cells(i, 10).Value > 0 Then
         ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    
        
        
    Next i
    'Print Values
        ws.Cells(2, 16).Value = ws.Cells(Max_Stock, 9).Value
        ws.Cells(2, 17).Value = Max_Percent
        ws.Cells(3, 16).Value = ws.Cells(Min_Stock, 9).Value
        ws.Cells(3, 17).Value = Min_Percent
        ws.Cells(4, 16).Value = ws.Cells(Vol_stock, 9).Value
        ws.Cells(4, 17).Value = Max_Volume
        ws.Range("Q2:Q3").NumberFormat = "#.##%"
    'Autofit all
    ws.Columns("A:Q").AutoFit
    
    ws.Columns("K").NumberFormat = "#.##%"
    
Next ws

End Sub

