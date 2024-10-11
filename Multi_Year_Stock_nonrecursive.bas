Attribute VB_Name = "Module1"
Sub Multi_Year_Stock()
    
    'Column Titles
    Range("I1, P1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q1").Value = "Value"
    
    'Row Titles
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
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
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To rowCount
    
    'If for ticker changes
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    total_vol = total_vol + Cells(i, 7).Value
        
        'What if there is a zero total volume
        If total_vol = 0 Then
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
        Else
        'Non-zero staring value
            If Cells(start_val, 3) = 0 Then
                For new_value = start_val To i
                    If Cells(new_value, 3).Value <> 0 Then
                        start_val = new_value
                        Exit For
                    End If
                Next new_value
                
            End If
    
            'Quarterly and Percent Changes
            quart_change = (Cells(i, 6) - Cells(start_val, 3))
            percent_change = quart_change / Cells(start_val, 3)
    
            'Next ticker
            start_val = i + 1
    
            'The results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Dim k As Integer
            For k = 2 To 2001
                If Cells(k, 10).Value > 0 Then
                    Cells(k, 10).Interior.ColorIndex = 4
                ElseIf Cells(k, 10).Value < 0 Then
                    Cells(k, 10).Interior.ColorIndex = 3
                Else
                    Cells(k, 10).Interior.ColorIndex = 0
                End If
            Next k
    
            Range("J" & 2 + j).Value = quart_change
            Range("J" & 2 + j).NumberFormat = "0.00"
    
            Range("K" & 2 + j).Value = percent_change
            Range("K" & 2 + j).NumberFormat = "0.00%"
    
            Range("L" & 2 + j).Value = total_vol
            
        End If
        
        'Variables for new ticker
        total_vol = 0
        quart_change = 0
        j = j + 1
        
        Else
            total_vol = total_vol + Cells(i, 7).Value
        End If
        
    Next i
    
    'Finding the greatest values
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))
    
    'Match Function
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)
    
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)
    
                 
End Sub
