Sub SummaryStocks()
Dim currentTicker As String
Dim nextTicker As String
Dim totalrows As LongLong
Dim total As LongLong
Dim summary_row As LongLong
Dim openprice As Single
Dim closeprice As Single
Dim yearlychange As LongLong



totalrows = Cells(Rows.Count, "A").End(xlUp).Row
summary_row = 2
For currentrow = 2 To totalrows
    currentTicker = Cells(currentrow, 1).Value
    nextTicker = Cells(currentrow + 1, 1).Value
    total = total + Cells(currentrow, 7).Value
    If currentrow = 2 Then
        openprice = Cells(currentrow, 3).Value
    End If
      
    
    If currentTicker <> nextTicker Then
        closeprice = Cells(currentrow, 6).Value
        
        
        Cells(summary_row, 9).Value = currentTicker
        Cells(summary_row, 12).Value = total
        
        Cells(summary_row, 10).NumberFormat = "0.00"
        
        Cells(summary_row, 10).Value = Round(closeprice - openprice, 2)
        If Cells(summary_row, 10).Value <= 0 Then
                    Cells(summary_row, 10).Interior.ColorIndex = 3
                Else
                    Cells(summary_row, 10).Interior.ColorIndex = 4
                End If
        
        Cells(summary_row, 11).NumberFormat = "0.00%"
        
        If openprice <> 0 Then
            Cells(summary_row, 11).Value = Round(closeprice - openprice, 2) / openprice
        End If
        
        
        summary_row = summary_row + 1
        
        total = 0
        openprice = Cells(currentrow + 1, 3).Value
    
    End If
Next currentrow
End Sub


Sub VBAHomework()
    For Each Ws In Worksheets
        Ws.Activate
        Call CalculateSummary
    Next Ws
End Sub
Sub CalculateSummary()
   
    Debug.Print ActiveSheet.Name
    Call SetTitle
    SummaryStocks
End Sub
Sub SetTitle()
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("I:O").Columns.AutoFit
End Sub
