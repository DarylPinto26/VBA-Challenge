Attribute VB_Name = "Module1"
Sub VBA_Challenge_2()


Dim Ticker As String
Dim open_year As Double
Dim close_year As Double
Dim yearly_change As Double
Dim total_stock_volume As Double
Dim percent_change As Double
Dim start_row As Integer
Dim ws As Worksheet

'First summary dataset
For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    start_row = 2
    previous_i = 1
    total_stock_volume = 0
    Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            previous_i = previous_i + 1
            open_year = ws.Cells(previous_i, 3).Value
            close_year = ws.Cells(i, 6).Value
            
            For j = previous_i To i
                total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
            Next j
            
            If open_year = 0 Then
                percent_change = close_year
                
            Else
                yearly_change = close_year - open_year
                percent_change = yearly_change / open_year
                
                
            End If
                
            ws.Cells(start_row, 9).Value = Ticker
            ws.Cells(start_row, 10).Value = yearly_change
            ws.Cells(start_row, 11).Value = percent_change
            ws.Cells(start_row, 11).NumberFormat = "0.00%"
            ws.Cells(start_row, 12).Value = total_stock_volume
            
            start_row = start_row + 1
            total_stock_volume = 0
            yearly_change = 0
            percent_change = 0
            previous_i = i
            End If
            
        Next i

    'Additional Functionality
    kLastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    Increase = 0
    Decrease = 0
    Greatest = 0
        
        For k = 3 To kLastRow
        
            last_k = k - 1
            current_k = ws.Cells(k, 11).Value
            previous_k = ws.Cells(last_k, 11).Value
            Volume = ws.Cells(k, 12).Value
            Previous_volume = ws.Cells(last_k, 12).Value
        
            If Increase > current_k And Increase > previous_k Then
                Increase = Increase
                increase_name = ws.Cells(k, 9).Value
                
            ElseIf current_k > Increase And current_k > previous_k Then
                Increase = current_k
                increase_name = ws.Cells(k, 9).Value
                
            ElseIf previous_k > Increase And previous_k > current_k Then
                Increase = previous_k
                increase_name = ws.Cells(last_k, 9).Value
            
            End If
            
            If Decrease < current_k And Decrease < previous_k Then
                Decrease = Decrease
                decrease_name = ws.Cells(k, 9).Value
            
            ElseIf current_k > Increase And current_k < previous_k Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value
            
            ElseIf previous_k < Increase And previous_k < current_k Then
                Decrease = previous_k
                decrease_name = ws.Cells(last_k, 9).Value
            
            End If
            
            If Greatest > Volume And Greatest > Previous_volume Then
                Greatest = Greatest
                greatest_name = ws.Cells(k, 9).Value
                
                
            ElseIf Volume > Greatest And Volume > Previous_volume Then
                Greatest = Volume
                greatest_name = ws.Cells(k, 9).Value
            
            ElseIf Previous_volume > Greatest And Previous_volume > Volume Then
                Greatest = Previous_volume
                greatest_name = ws.Cells(last_k, 9).Value
            
            End If
            
        Next k
                         
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("N1").Value = "Additional Functionality"
    
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest
    
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    
    'Conditional Formatting for yearly change
    jLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For l = 2 To jLastRow
            If ws.Cells(l, 10) > 0 Then
                ws.Cells(l, 10).Interior.ColorIndex = 4
        
            Else
                ws.Cells(l, 10).Interior.ColorIndex = 3
            End If
            
        Next l
    
Next ws
    
End Sub
