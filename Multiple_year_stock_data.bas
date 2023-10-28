Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

For Each ws In Worksheets

Dim Ticker As String
Dim vol As Double
    vol = 0
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
Dim open_price As Double
    open_price = ws.Cells(2, 3).Value
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("L" & Summary_Table_Row).Value = vol
            close_price = ws.Cells(i, 6).Value
            yearly_change = (close_price - open_price)
            Range("J" & Summary_Table_Row).Value = yearly_change
                
                If (open_price = 0) Then
                    percent_change = 0
                    
                Else
                    percent_change = yearly_change / open_price
                
                End If
            
            ws.Range("k" & Summary_Table_Row).Value = percent_change
            ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
            
            Summary_Table_Row = Summary_Table_Row + 1
            vol = 0
            open_price = ws.Cells(i + 1, 3)
        
        
        
        Else
            vol = vol + ws.Cells(i, 7).Value
            
        End If
    Next i
    
lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrow_summary_table
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

    For i = 2 To lastrow_summary_table
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
    Next ws
    
End Sub
