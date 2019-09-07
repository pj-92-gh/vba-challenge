Sub vba_stock()

Dim ws As Worksheet

For Each ws In Worksheets

    ' Setting headers for VBA scripts '
    
    ws.Range("J1").Value = "<ticker>"
    ws.Range("K1").Value = "<open>"
    ws.Range("L1").Value = "<close>"
    ws.Range("M1").Value = "<difference>"
    ws.Range("N1").Value = "<% diff>"
    ws.Range("O1").Value = "<tot volume>"

    ' Setting up counter variable - not necessary but found it useful '
    
    Dim count_hold As Integer
    
    ' Now, we start setting up the initial loop '

    tick_range = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' The first count_hold variable holds the counter to paste the ticker name/year start/year end/different/% difference '
    ' The row_count variable is to find out the first date for any given ticker name. Then, we can store the year start price '
    ' The sum row counter is used to track volume totals '
    
    count_hold = 2
    row_count = 1
    sum_row = 0
    
    ' Main loop begins '
    ' Checks for unique ticker names and then adds/grabs values as needed '
    
        For y = 2 To tick_range
            If ws.Cells(y, 1).Value = ws.Cells(y + 1, 1).Value And row_count = 1 Then
                ws.Cells(count_hold, 11).Value = ws.Cells(y, 3).Value
                year_start_price = ws.Cells(y, 3).Value
                row_count = row_count + 1
                sum_row = ws.Cells(y, 7).Value
            ElseIf ws.Cells(y, 1).Value = ws.Cells(y + 1, 1).Value Then
                sum_row = sum_row + ws.Cells(y, 7).Value
            ElseIf ws.Cells(y + 1, 1).Value <> ws.Cells(y, 1).Value And year_start_price <> 0 Then
                ws.Cells(count_hold, 10).Value = ws.Cells(y, 1).Value
                ws.Cells(count_hold, 12).Value = ws.Cells(y, 6)
                year_end_price = ws.Cells(y, 6)
                difference = (year_end_price - year_start_price)
                ws.Cells(count_hold, 13).Value = difference
                ws.Cells(count_hold, 14).Value = Str((((year_end_price - year_start_price) / (year_start_price)) * 100)) + "%"
                sum_row = sum_row + ws.Cells(y, 7).Value
                ws.Cells(count_hold, 15).Value = sum_row
                count_hold = count_hold + 1
                row_count = 1
            End If
        Next y
    
    ' Summary information set up '
    
    ws.Range("Q2").Value = "Greatest % Decrease"
    ws.Range("Q3").Value = "Greatest % Increase"
    ws.Range("Q4").Value = "Greatest Total Volume"
    ws.Range("R1").Value = "Tick"
    ws.Range("S1").Value = "Value"
    
    min_num = 0
    max_num = 0
    tot_max = 0
    
    For q = 2 To ws.Cells(Rows.Count, 14).End(xlUp).Row
        If min_num > ws.Cells(q, 14).Value Then
            min_num = ws.Cells(q, 14).Value
            min_tick = ws.Cells(q, 10).Value
        End If
        If max_num < ws.Cells(q, 14).Value Then
            max_num = ws.Cells(q, 14).Value
            max_tick = ws.Cells(q, 10).Value
        End If
        If tot_max < ws.Cells(q, 15).Value Then
            tot_max = ws.Cells(q, 15).Value
            tot_tick = ws.Cells(q, 10).Value
        End If
        ws.Range("S2").Value = Str(min_num * 100) + "%"
        ws.Range("R2").Value = min_tick
        ws.Range("S3").Value = Str(max_num * 100) + "%"
        ws.Range("R3").Value = max_tick
        ws.Range("S4").Value = tot_max
        ws.Range("R4").Value = tot_tick
    Next q
    
    ' Conditional loop logic '
    
    For O = 2 To ws.Cells(Rows.Count, 13).End(xlUp).Row
        If ws.Cells(O, 13).Value < 0 Then
            ws.Range("M" & O).Interior.ColorIndex = 3
        Else
            ws.Range("M" & O).Interior.ColorIndex = 4
        End If
    Next O

Next ws

End Sub
