Sub StockData():
    
    ' Set worksheets in the book:
    Dim ws As Worksheet
    
    For Each ws In Worksheets

    ' Set Ticker Symbol, Total Volume
    Dim Ticker As String
    Dim TotalVolume As Double
    ' Set Total Volume to 0
    TotalVolume = 0
    
    ' Set Open Price and Start Value
    Dim OpenPrice As Double
        ' Set first ticker's Open Price to the value in Cells (2,3)
        OpenPrice = ws.Cells(2, 3).Value
      

    ' Set Summary table: 1) Title and Tracking starting row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
                       ' 2) Set summary table starting on 2nd row (first row for titles)
        Dim SummaryRow As Double
            SummaryRow = 2

    ' Find the last row number in the Raw Data
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Let's LOOP!
            Dim i As Long
            For i = 2 To LastRow
                ' Change Cell Test: when the next cell <> the current cell, indicating ticker symbol will change in the next cell
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                  
                ' Set the Ticker symbol and Volume number
               Ticker = ws.Cells(i, 1).Value
               TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            ' Set Close Price
                 Dim ClosePrice As Double
                 ClosePrice = ws.Cells(i, 6).Value
                 
                 
            ' Print results into summary table
            ws.Range("I" & SummaryRow).Value = Ticker            'Ticker Symbol
            ws.Range("L" & SummaryRow).Value = TotalVolume       'Total Volume per Ticker
             
            ' Set Yearly Change and Percent Change Calculations
                 Dim YearlyChange As Double
                     YearlyChange = ClosePrice - OpenPrice
                     
                 Dim PercentChange As Double
                 If (OpenPrice = 0 And ClosePrice = 0) Then
                      PercentChange = 0
                    ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                        PercentChange = 1
                        Else
                           PercentChange = (YearlyChange / OpenPrice)
                  End If
      
            ws.Range("J" & SummaryRow).Value = YearlyChange     'Yearly Change per Ticker
            ws.Range("K" & SummaryRow).Value = PercentChange    'Change Percentage per Ticker
            'Change value format to %
            ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
        
             
                ' Reset Summary Row to the next row for the next record
                SummaryRow = SummaryRow + 1
                ' Reset TotalVolume back to 0
                TotalVolume = 0
                ' Reset OpenPrice
                OpenPrice = ws.Cells(i + 1, 3).Value

                
            ' If cell value is not changing, simply add volumes together by each row
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
               End If
                  
         Next i


    ' Moderate: Apply Colors in summary table
    ' Find the last row number in the Summary Table
        Dim SUMLastRow As Double
        SUMLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        ' Loop for Colors
        Dim j As Long
        For j = 2 To SUMLastRow
            If ws.Cells(j, 10).Value >= 0 Then
                 ws.Cells(j, 10).Interior.ColorIndex = 4
                
               Else
              
                 ws.Cells(j, 10).Interior.ColorIndex = 3

            End If
        
        Next j
        

    ' Hard: Define "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"
    ' Set separate table

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

       ' Loop thru summary table looking for Greatest Increase and Greatest Decrease
             For z = 2 To SUMLastRow
       
            If ws.Cells(z, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & SUMLastRow)) Then
                  ws.Range("P2").Value = ws.Cells(z, 9).Value
                  ws.Range("Q2").Value = ws.Cells(z, 11).Value
                  ws.Range("Q2").NumberFormat = "0.00%"
                
                ElseIf ws.Cells(z, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & SUMLastRow)) Then
                          ws.Range("P3").Value = ws.Cells(z, 9).Value
                          ws.Range("Q3").Value = ws.Cells(z, 11).Value
                          ws.Range("Q3").NumberFormat = "0.00%"
                
                    ElseIf ws.Cells(z, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & SUMLastRow)) Then
                              ws.Range("P4").Value = ws.Cells(z, 9).Value
                              ws.Range("Q4").Value = ws.Cells(z, 12).Value
          
          End If
         
        Next z

    Next ws

End Sub