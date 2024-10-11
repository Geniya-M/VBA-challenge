Attribute VB_Name = "Module1"
Sub Multi_Year_Stock()
    
    'Loop through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    'Column Titles
    ws.Range("I1, P1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Value"
    
    'Row Titles
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Set variables and values
    Dim rowCount As Long
    Dim total_vol As Double
    Dim quart_change As Double
    Dim i As Long
    Dim j As Integer
    Dim start_val As Long
    
    total_vol = 0
    quart_change = 0
    start_val = 2
    j = 0
    
    Dim daily_change As Double
    Dim percent_change As Double
    Dim ave_chang As Double
    
    'Find the last row
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To rowCount
    
    'If for ticker changes
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    total_vol = total_vol + ws.Cells(i, 7).Value
        
        'What if there is a zero total volume
        If total_vol = 0 Then
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = 0
            ws.Range("K" & 2 + j).Value = "%" & 0
            ws.Range("L" & 2 + j).Value = 0
        Else
        'Non-zero staring value
            If ws.Cells(start_val, 3) = 0 Then
                For new_value = start_val To i
                    If ws.Cells(new_value, 3).Value <> 0 Then
                        start_val = new_value
                        Exit For
                    End If
                Next new_value
                
            End If
    
            'Quarterly and Percent Changes
            quart_change = (ws.Cells(i, 6) - ws.Cells(start_val, 3))
            percent_change = quart_change / ws.Cells(start_val, 3)
    
            'Next ticker
            start_val = i + 1
    
            'The results
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            Dim k As Integer
            For k = 2 To 2001
                If ws.Cells(k, 10).Value > 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(k, 10).Value < 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 0
                End If
            Next k
    
            ws.Range("J" & 2 + j).Value = quart_change
            ws.Range("J" & 2 + j).NumberFormat = "0.00"
    
            ws.Range("K" & 2 + j).Value = percent_change
            ws.Range("K" & 2 + j).NumberFormat = "0.00%"
    
            ws.Range("L" & 2 + j).Value = total_vol
            
        End If
        
        'Variables for new ticker
        total_vol = 0
        quart_change = 0
        j = j + 1
        
        Else
            total_vol = total_vol + ws.Cells(i, 7).Value
        End If
        
    Next i
    
    'Finding the greatest values
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
    
    'Match Function
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
    
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)
    
Next ws
                 
End Sub
