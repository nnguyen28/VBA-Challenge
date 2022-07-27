Sub VBA_Challenge():

'Insert data
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

'Create variable
Dim ticker As String
    
Dim yearopen As Double
    yearopen = 0
    
Dim yearclose As Double
    yearclose = 0
    
Dim summary_table_index As Integer
    summary_table_index = 2
    
Dim vol As Double
    vol = 0

'Loop the data
    For i = 2 To 22771

        vol = vol + Cells(i, 7).Value
        yearclose = yearclose + Cells(i, 6).Value
        yearopen = yearopen + Cells(i, 3).Value
        
        'Estimate yearly change of the stock
        yearly_change = Cells(i, 6).Value - Cells(i, 3).Value
            Range("j" & summary_table_index).Value = yearly_change
            
        'Estimate percent change of the stock
        percent_change = yearly_change / yearopen
            Range("K" & summary_table_index).Value = percent_change
                
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            'Input the ticker in ticker column
            ticker = Cells(i, 1).Value
            Range("i" & summary_table_index).Value = ticker
            
            'Sum the total amount and put it into volume column
            Range("L" & summary_table_index).Value = vol
            
            summary_table_index = summary_table_index + 1
            
        vol = 0
        
        End If
    Next i

End Sub