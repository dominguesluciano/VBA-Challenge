Sub stock_analysis():
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ' set variables
    Dim total_stock As Double
    Dim I As Long
    Dim diff As Double
    Dim summary_table As Integer
    Dim start_row As Long
    Dim row_count As Long
    Dim percent_diff As Double
    Dim averageChange As Double
    Dim Total As Double
    Dim greatest As Double
    Dim smallest As Double
    Dim gstock As Single
    
    

    ' Set header for each column
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set initial values
    summary_table = 2
    total_stock = 0
    diff = 0
    start_row = 2
    Total = 0
    find_value = 0
        
    ' get the row number of the last row with data
    row_count = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For I = 2 To row_count
        ' If ticker is different then previous one then print results
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            ' Stores results in variables
            Total = total_stock + ws.Cells(I, 7).Value
            ' Handle zero total_stock volume
            If Total = 0 Then
                ' print results in summary table
                ws.Range("I" & summary_table).Value = ws.Cells(I, 1).Value
                ws.Range("J" & summary_table).Value = 0
                ws.Range("K" & summary_table).Value = "%" & 0
                ws.Range("L" & summary_table).Value = 0
            Else
                ' Find first non zero value
                If ws.Cells(start_row, 3) = 0 Then
                    For find_value = start_row To I
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start_row = find_value
                            
                            Exit For
                        End If
                     Next find_value
                End If
                
            
                 
                ' Calculate difference
                diff = (ws.Cells(I, 6) - ws.Cells(start_row, 3))
                percent_diff = Round((diff / ws.Cells(start_row, 3) * 100), 2)
                
                
                ' start of the next stock ticker
                start_row = I + 1
                ' print results
                ws.Range("I" & summary_table).Value = ws.Cells(I, 1).Value
                ws.Range("J" & summary_table).Value = Round(diff, 2)
                ws.Range("K" & summary_table).Value = "%" & percent_diff
                ws.Range("L" & summary_table).Value = Total
              
                If ws.Range("J" & summary_table) > 0 Then
                    ws.Range("J" & summary_table).Interior.ColorIndex = 4
                    
                ElseIf ws.Range("J" & summary_table) < 0 Then
                    ws.Range("J" & summary_table).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & summary_table).Interior.ColorIndex = 0
                End If
                
                
   
            End If
            ' reset variables for new stock ticker
            total_stock = 0
            diff = 0
            summary_table = summary_table + 1
        ' If ticker is still the same add results
        Else
            total_stock = total_stock + ws.Cells(I, 7).Value
        End If

        
    Next I
    greatest = WorksheetFunction.Max(ws.Range("K2:K" & ws.Cells(Rows.Count, "K").End(xlUp).Row))
    smallest = WorksheetFunction.Min(ws.Range("K2:K" & ws.Cells(Rows.Count, "K").End(xlUp).Row))
    gstock = WorksheetFunction.Max(ws.Range("L2:L" & ws.Cells(Rows.Count, "L").End(xlUp).Row))
    
    'print headers for greatest numbers
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    
    'print row labels for greatest numbers
    ws.Range("P2").Value = "Greatest % increase"
    ws.Range("P3").Value = "Greatest % decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    
    'print return of variables for greatest numbers
    ws.Range("R2").Value = "%" & greatest * 100
    ws.Range("R3").Value = "%" & smallest * 100
    ws.Range("R4").Value = gstock

    
    Next ws



    
End Sub












