Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

    'create a worksheet loop
    For Each ws In Worksheets
    
        'create headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'create a loop for the data
        Dim i, j, tableRow As Long
        Dim openYear, closeYear As Double
        Dim volume As Double
        
       'set dimensions for summary table
        Dim ticker As String
        Dim yearlyChange, percentChange As Double
        
        'set initial values for summary table
        tableRow = 2
        volume = 0
        openYear = 0
        closeYear = 0
        
        'define last row for data
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
          
        For i = 2 To lastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'extract data if ticker not the same as prior row
                closeYear = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
                
                'assign value to headers and paste to summary table
                ticker = ws.Cells(i, 1).Value
                ws.Cells(tableRow, 9).Value = ticker
                
                yearlyChange = closeYear - openYear
                ws.Cells(tableRow, 10).Value = yearlyChange
                
                'format to two decimals
                ws.Cells(tableRow, 10).NumberFormat = "0.00"
                
                    'set colour condition for yearly change
                    If yearlyChange < 0 Then
                    
                        ws.Cells(tableRow, 10).Interior.ColorIndex = 3
                        Else: ws.Cells(tableRow, 10).Interior.ColorIndex = 4
                        
                    End If
                
                'continue pasting values to summary table
                percentChange = yearlyChange / openYear
                ws.Cells(tableRow, 11).Value = percentChange
                
                    'format cell to percent
                    ws.Cells(tableRow, 11).NumberFormat = "0.00%"
                    
                'continue pasting values to summary table
                ws.Cells(tableRow, 12).Value = volume
                
                'prepare values for next summary table row
                tableRow = tableRow + 1
                volume = 0
            
                'if ticker is the same as prior row. Check for open year date
                ElseIf ws.Cells(i, 2).Value = ws.Name & "0102" Then
                openYear = ws.Cells(i, 3).Value
                volume = volume + ws.Cells(i, 7).Value
                
                'if not open year date
                Else: volume = volume + ws.Cells(i, 7).Value
                
            End If
         
        Next i
        
        'autofit columns
        ws.Range("I:L").Columns.AutoFit
        
        '_____________________________________________________________________
        'BONUS SECTION
        '_____________________________________________________________________
        
        'create headers for bonus table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'create loop to search summary table
        lastRowSummary = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        Dim increaseTicker, decreaseTicker, totalTicker As String
        Dim greatestIncrease, greatestDecrease, greatestTotal As Double
        
        'store values
        increaseTicker = ws.Range("P2").Value
        decreaseTicker = ws.Range("P3").Value
        totalTicker = ws.Range("P4").Value
        greatestIncrease = ws.Range("Q2").Value
        greatestDecrease = ws.Range("Q3").Value
        greatestTotal = ws.Range("Q4").Value
        
        'assign starting values
        ws.Range("P2") = ws.Cells(2, 9).Value
        ws.Range("Q2") = ws.Cells(2, 11).Value
        ws.Range("P3") = ws.Cells(2, 9).Value
        ws.Range("Q3") = ws.Cells(2, 11).Value
        ws.Range("P4") = ws.Cells(2, 9).Value
        ws.Range("Q4") = ws.Cells(2, 12).Value
        
        'format cells
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'check greatest % increase
        For i = 2 To lastRowSummary
        
            If ws.Cells(i + 1, 11).Value > greatestIncrease And ws.Cells(i + 1, 11).Value <> "" Then
            
                greatestIncrease = ws.Cells(i + 1, 11).Value
                increaseTicker = ws.Cells(i + 1, 9).Value
                
                'paste values to table
                ws.Range("P2").Value = increaseTicker
                ws.Range("Q2").Value = greatestIncrease
                
                'format cell to percent
                ws.Range("Q2").NumberFormat = "0.00%"
        
                Else
                
            End If
    
        Next i
        
 '_______________
 
        'check greatest % decrease
        For i = 2 To lastRowSummary
        
            If ws.Cells(i + 1, 11).Value < ws.Range("Q3").Value And ws.Cells(i + 1, 11).Value <> "" Then
            
                greatestDecrease = ws.Cells(i + 1, 11).Value
                decreaseTicker = ws.Cells(i + 1, 9).Value
                
                'paste values to table
                ws.Range("P3").Value = decreaseTicker
                ws.Range("Q3").Value = greatestDecrease
                
                'format cell to percent
                ws.Range("Q3").NumberFormat = "0.00%"
                
                Else
                
            End If
    
        Next i

 '_________________

         'check greatest volume
        For i = 2 To lastRowSummary
        
            If ws.Cells(i + 1, 12).Value > greatestTotal And ws.Cells(i + 1, 12).Value <> "" Then
                
                greatestTotal = ws.Cells(i + 1, 12).Value
                totalTicker = ws.Cells(i + 1, 9).Value
                
                'paste values to table
                ws.Range("P4").Value = totalTicker
                ws.Range("Q4").Value = greatestTotal
                
                Else
                
            End If
    
        Next i

 '_________________
            
    'autofit columns
    ws.Range("O:Q").Columns.AutoFit
            
    Next ws
    
End Sub
