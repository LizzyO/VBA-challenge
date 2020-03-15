Sub LOStock()

'run on every worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
  ' Set initial variables
Dim Ticker_name As String
'set open price from the start
Dim Open_price As Double
    Open_price = Cells(2, 3).Value
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock As Double
    Total_Stock = 0
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'header for each column
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"

  'Challenge state values
  Cells(2, 14).Value = "Greatest % Increase"
  Cells(3, 14).Value = "Greatest % Decrease"
  Cells(4, 14).Value = "Greatest Total Volume"
  Cells(1, 15).Value = "Ticker"
  Cells(1, 16).Value = "Value"
  
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastrow
     ' Check if we are still within the same ticker, if it is not...
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Set ticker name
            Ticker_name = Cells(i, 1).Value
            
            'set close price
            Close_Price = Cells(i, 6).Value
            
            'calculate yearly change
            Yearly_Change = Close_Price - Open_price
            
            'calculate percent change
                    If Open_price = 0 Then
                    Percent_Change = 0
                    Else
                    Percent_Change = Yearly_Change / Open_price
                    End If
            'calculate total stock
            Total_Stock = Total_Stock + Cells(i, 7).Value

                    ' Print in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker_name
            Range("j" & Summary_Table_Row).Value = Yearly_Change
            Range("k" & Summary_Table_Row).Value = Percent_Change
            Range("l" & Summary_Table_Row).Value = Total_Stock
                    
                    If Yearly_Change >= 0 Then
                        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    ElseIf Yearly_Change < 0 Then
                        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    End If
                         
            ' For the next loop: row to the summary table row, set new open price and reset total stock
            Summary_Table_Row = Summary_Table_Row + 1
            Open_price = Cells(i + 1, 3).Value
            Total_Stock = 0
            
         ' If the cell immediately following a row is the same ticker...
            Else
            ' Add to the Stock Total
            Total_Stock = Total_Stock + Cells(i, 7).Value
            End If
        Next i
        Columns("k").NumberFormat = "0.00%"
        
        'loop throw the new ticker list range
        tlastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        For j = 2 To tlastrow
        'find greatest % row and format as percent
        If Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:k" & tlastrow)) Then
                Cells(2, 15).Value = Cells(j, 9).Value
                Cells(2, 16).Value = Cells(j, 11).Value
                Cells(2, 16).NumberFormat = "0.00%"
            'find least % row and format as percent
            ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(ws.Range("k2:k" & tlastrow)) Then
                 Cells(3, 15).Value = Cells(j, 9).Value
                 Cells(3, 16).Value = Cells(j, 11).Value
                 Cells(3, 16).NumberFormat = "0.00%"
            'find greatest volume row and format as percent  
            ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:l" & tlastrow)) Then
                 Cells(4, 15).Value = Cells(j, 9).Value
                 Cells(4, 16).Value = Cells(j, 12).Value
            End If
        Next j
    Next ws
End Sub

