Attribute VB_Name = "Module2"
Sub ticker()
    ' Looping through worksheets
    For Each ws In Worksheets 'Use ThisWorkbook to reference the workbook containing the macro
        ' Define worksheet name WorksheetName
        ws.Activate
        WorksheetName = ws.Name
        ' Name column header for analysis criteria
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(1, "O").Value = "Ticker"
        ws.Cells(1, "P").Value = "Value"
       
        
        ' Initial conditions for ticker summary
        initialStartRow = 2
        'r = 2 ' Start from row 2
        TotalStockVolume = 0
        OpenPrice = Cells(2, "C").Value
        ' Define last row
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Loop through rows
        For r = 2 To LastRow
        
            ' Check if the ticker name changes
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
                TotalStockVolume = TotalStockVolume + ws.Cells(r, "G")
                ' Define Ticker name
                TickerName = ws.Cells(r, "A").Value
                ' Yearly end "<close>"
                ClosePrice = ws.Cells(r, "F").Value
                ' Year change
                YearChange = ClosePrice - OpenPrice
                ' Percent change
               PercentChange = YearChange / OpenPrice * 100
               
                                   
                ' Ticker Name in Summary Table
               ws.Cells(initialStartRow, "I").Value = TickerName
                ' Yearly Change in summary table
                ws.Cells(initialStartRow, "J").Value = YearChange
              
                ' Percent change in summary table
               ws.Cells(initialStartRow, "K").Value = "%" & PercentChange
                ' Stock Volume row
                ws.Cells(initialStartRow, "L").Value = TotalStockVolume
                
                If YearChange > 0 Then
                ws.Cells(initialStartRow, "J").Interior.ColorIndex = 4
                Else
                ws.Cells(initialStartRow, "J").Interior.ColorIndex = 3
                End If
               
               
                ' Total Stock Volume in summary table
                TotalStockVolume = 0
                ' Reset initialStartRow
                initialStartRow = initialStartRow + 1
                'Format Percent row with % sign (K2 isnt working?)
                
                OpenPrice = Cells(r + 1, "C").Value
            Else
                 ' Add Stock Volume for Each Ticker
                TotalStockVolume = TotalStockVolume + ws.Cells(r, "G")
                
                
            End If
            
     
            
        Next r
        
        Cells(1, "O").Value = "Ticker"
        Cells(1, "P").Value = "Value"
        Cells(2, "N").Value = "Greatest % Increase"
        Cells(3, "N").Value = "Greatest % Decrease"
        Cells(4, "N").Value = "Greatest Total Volume"
       summary_count = Cells(Rows.Count, "K").End(xlUp).Row
        'LastRowPT = ws.Cells(ws.Rows.Count, "I").End(x1Up).Row
        'MsgBox (LastRowPT)
        
        'define initial start condition
        PerIncrease = 0
        PerDecrease = 0
        GrTotVol = 0
        
        For i = 2 To summary_count
        
            If Cells(i, "K").Value > PerIncrease Then
                PerIncrease = Cells(i, "K").Value
                TickName = Cells(i, "I").Value
            End If
                   
        Next i
        
        ws.Cells(2, "O").Value = TickName
        ws.Cells(2, "P").Value = PerIncrease
        
        For i = 2 To summary_count
        
            If Cells(i, "K").Value < PerDecrease Then
                PerDecrease = Cells(i, "K").Value
                TickName = Cells(i, "I").Value
            End If
                   
        Next i
        
        ws.Cells(3, "O").Value = TickName
        ws.Cells(3, "P").Value = PerDecrease
        
        For i = 2 To summary_count
        
            If Cells(i, "L").Value > GrTotVol Then
                GrTotVol = Cells(i, "L").Value
                TickName = Cells(i, "I").Value
            End If
                   
        Next i
        
        ws.Cells(4, "O").Value = TickName
        ws.Cells(4, "P").Value = GrTotVol
        
      Columns("A:P").AutoFit
      
    Next ws
    
    MsgBox ("Finish")
End Sub


