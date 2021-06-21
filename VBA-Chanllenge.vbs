Sub stock_summary_2()
Dim ws As Worksheet
Dim lastrow, lastrow_summary As LongLong
Dim index As Integer

For index = 1 To ActiveWorkbook.Worksheets.Count
    Set ws = Worksheets(index)
    With ws
    
        .Cells(1, "I").Value = "Ticker"
        .Cells(1, "J").Value = "Yearly Change"
        .Cells(1, "K").Value = "Percent Change"
        .Cells(1, "L").Value = "Total Stock Volumn"
        .Cells(2, "O").Value = "Great % Increase"
        .Cells(3, "O").Value = "Great % Decrease"
        .Cells(4, "O").Value = "Great Total Volumn"
        .Cells(1, "P").Value = "Ticker"
        .Cells(1, "Q").Value = "Value"
    
        '''Dim lastrow, lastrow_summary As LongLong
    
        summary_table_row = 2
        opening = 2
        total = 0
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        lastrow_summary = ws.Cells(Rows.Count, "A").End(xlUp).Row
         For i = 2 To lastrow
            If .Cells(i + 1, "A").Value <> .Cells(i, "A").Value Then
                ticker = .Cells(i, "A").Value
                change_dollar = .Cells(i, "F").Value - .Cells(opening, "C").Value
                If .Cells(opening, "C").Value <> 0 Then
                    change_p = change_dollar / .Cells(opening, "C").Value
                Else
                    change_p = 0
                End If
                total = total + .Cells(i, "G").Value
                .Cells(summary_table_row, "I").Value = ticker
                .Cells(summary_table_row, "J").Value = change_dollar
                .Cells(summary_table_row, "K").Value = change_p
                .Cells(summary_table_row, "L").Value = total
                total = 0
                summary_table_row = summary_table_row + 1
                opening = i + 1
            Else
                total = total + .Cells(i, "G").Value
            End If
        Next i
    
        For j = 2 To lastrow_summary
            If .Cells(j, "J").Value < 0 Then
                .Cells(j, "J").Interior.ColorIndex = 3
            ElseIf .Cells(j, "J").Value > 0 Then
                .Cells(j, "J").Interior.ColorIndex = 4
            End If
        Next j
    
        .Range("K2" & ":K" & lastrow_summary).NumberFormat = "0.00%"
    
        Max = 0
        Min = 0
        Max_v = 0
    
        For k = 2 To lastrow_summary
            If .Cells(k, "K").Value > Max Then
                Max = .Cells(k, "K").Value
                max_stock = .Cells(k, "I").Value
            End If
        Next k
            .Cells(2, "P").Value = max_stock
            .Cells(2, "Q").Value = Max
            .Cells(2, "Q").NumberFormat = "0.00%"
        
        For k = 2 To lastrow_summary
            If .Cells(k, "K").Value < Min Then
                Min = .Cells(k, "K").Value
                min_stock = .Cells(k, "I").Value
            End If
        Next k
            .Cells(3, "P").Value = min_stock
            .Cells(3, "Q").Value = Min
            .Cells(3, "Q").NumberFormat = "0.00%"
        
        For k = 2 To lastrow_summary
            If .Cells(k, "L").Value > Max_v Then
                Max_v = .Cells(k, "L").Value
                Max_v_stock = .Cells(k, "I").Value
            End If
        Next k
        .Cells(4, "P") = Max_v_stock
        .Cells(4, "Q").Value = Max_v
    End With
Next index
End Sub