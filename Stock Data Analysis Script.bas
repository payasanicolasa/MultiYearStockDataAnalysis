Attribute VB_Name = "Module1"
Sub StockChallenge():

'Adjust script to run on every worksheet in the doc (every year).
'NOTE: Make sure the script acts the same on every sheet
  For Each ws In Worksheets
      ws.Activate

    Dim i, lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim Row As Integer
    Row = 0
    
    Dim OpenPrice, ClosePrice, YChange, PChange As Double
    OpenPrice = Cells(2, 3).Value
    
    Dim MaxP, MinP, MaxV, volume As Double
    MaxP = 0
    MinP = 0
    MaxV = 0
    volume = 0

    For i = 2 To lastrow
        'Consolidate Duplicate Tickers in New Column
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Row = Row + 1
            Cells(Row + 1, 9).Value = Cells(i, 1).Value
        
        'Calculate Total Volume
            volume = volume + Cells(i, 7).Value
            Cells(Row + 1, 12).Value = volume

        'Calculate Yearly Change (By Open/Close, Columns C & F)
        ClosePrice = Cells(i, 6).Value
        YChange = (ClosePrice - OpenPrice)
        Cells(Row + 1, 10).Value = YChange
             
        'Calculate Percent Change (By Date, Column B)
        If OpenPrice <> 0 Then
            PChange = YChange / OpenPrice
            Cells(Row + 1, 11).Value = PChange
            Cells(Row + 1, 11).Value = FormatPercent(Cells(Row + 1, 11))
        End If
                  
        'Use conditional formatting to highlight pos/neg Yearly Change in green/red
        If Cells(Row + 1, 10).Value < 0 Then
            Cells(Row + 1, 10).Interior.Color = vbRed
        ElseIf Cells(Row + 1, 10).Value > 0 Then
            Cells(Row + 1, 10).Interior.Color = vbGreen
        End If
      
        volume = 0
        OpenPrice = Cells(i + 1, 3)
        
    Else
        volume = volume + Cells(i, 7).Value
    End If
    
    Next i
            LastYChange = Cells(Rows.Count, 10).End(xlUp).Row
            For n = 2 To LastYChange
    
                'BONUS: Return the stock w the "Greatest % decrease" 'Return the stock w the "Greatest total volume"
                If MinP > Cells(n, 11).Value Then
                    MinP = Cells(n, 11).Value
                    Range("Q2").Value = MinP
                    Range("Q2").NumberFormat = "0.00%"
                    Range("P2").Value = Cells(n, 9).Value
                End If
                
                'BONUS: Return the stock w the "Greatest % increase"
                 If MaxP < Cells(n, 11).Value Then
                    MaxP = Cells(n, 11).Value
                    Range("Q3").Value = MaxP
                    Range("Q3").NumberFormat = "0.00%"
                    Range("P3").Value = Cells(n, 9).Value
                End If
                
                'BONUS: Return stock with Highest Total Volume
                If MaxV < Cells(n, 12).Value Then
                    MaxV = Cells(n, 12).Value
                    Range("Q4").Value = MaxV
                    Range("P4").Value = Cells(n, 9).Value
                    
                End If
            
            Next n
            
        'Create headers for new columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("O2").Value = "Greatest % Decrease"
        Range("O3").Value = "Greatest % Increase"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Autofit Column Width in all sheets
        ws.Cells.EntireColumn.AutoFit
  Next

End Sub
