Attribute VB_Name = "Module1"
Sub stockAnalysis():

    Dim total As Double
    Dim row As Long
    Dim rowCount As Long
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim summaryTableRow As Long
    Dim stockStartRow As Long
    
    'loop through all work sheets
    For Each ws In Worksheets
    
        ' Set title rows
       ws.Range("I1").Value = "Ticker"
       ws.Range("J1").Value = "Yearly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Stock Volume"
       ws.Range("P1").Value = "Ticker"
       ws.Range("Q1").Value = "Value"
       ws.Range("Q2").Value = "Greatest % Increase"
       ws.Range("Q3").Value = "Greatest % Decrease"
       ws.Range("Q4").Value = "Greatest Total Volume"
       
        
        'initialize the values
        summaryTableRow = 0
        total = 0
        yearlyChange = 0
        stocksStartRow = 2
        
        'get value for last row
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'loop till end of sheet
        For row = 2 To rowCount
            
            'check for changes column a
            If ws.Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
                Else
                    total = total + ws.Cells(row, 7).Value
                
                If total = 0 Then
                
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                
                Else
                    
                    If ws.Cells(stocksStartRow, 3).Value = 0 Then
                        For findValue = stocksStartRow To row
                        
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                stocksStartRow = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                    
                    yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(stocksStartRow, 3).Value)
                    percentChange = yearlyChange / ws.Cells(stocksStartRow, 3).Value
                    
                    ws.Range("I" & 2 + summaryTableRow).Value = Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
                    ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###"
                    
                    If yearlyChange > 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                    ElseIf yearlyChage < 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    End If
                 End If
                 
                 total = 0
                 yearlyChange = 0
                 summaryTableRow = summaryTableRow + 1
                 
                
                    
                    
            
            
            End If
            
        Next row
        
        'after looping through the rows, find the max and min and palce in Q2, Q3,Q4
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100 'value of the greatest increase
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100 'value of the greatest decrease
        ws.Range("Q4") = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100  'value of the greatest max volume
            
        'matching in order to match the ticker names with the values
        'tell the row in the summary table where the tickeratches the greatest increase
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
        
        'tell the row in the summary table where the tickeratches the greatest decrease
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P3").Value = ws.Cells(increaseNumber + 1, 9)
        
        'tell the row in the summary table where the tickeratches the greatest increase
        volumeNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P4").Value = ws.Cells(volumeNumber + 1, 9)
        
        
        
        ws.Columns("A:Q").AutoFit
        
        Next ws
End Sub
