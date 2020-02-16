Sub Stockdata()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
       
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim ticker_name As String
        Dim percent_change As Double
        Dim volume As Double
        volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set Initial Open Price
        open_price = Cells(2, Column + 2).Value
        
        For i = 2 To LastRow
         'Check if still the same ticker
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ticker_name = Cells(i, Column).Value
                ticker_name = Cells(Row, Column + 8).Value
                
                close_price = Cells(i, Column + 5).Value
                
                yearly_change = close_price - open_price
                Cells(Row, Column + 9).Value = yearly_change
                'Percent Change
                If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    percent_change = 1
                Else
                    percent_change = yearly_change / open_price
                    Cells(Row, Column + 10).Value = percent_change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
        
                volume = volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = volume
                Row = Row + 1
                open_price = Cells(i + 1, Column + 2)
                volume = 0
            Else
                volume = volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        'set colors
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
    'bonus part
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
    'find the greatest value and its ticker
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub

