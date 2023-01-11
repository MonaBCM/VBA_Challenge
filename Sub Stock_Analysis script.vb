
Sub StockAnalysis()

    Dim Ws As Worksheet
    For Each Ws In ActiveWorkbook.worksheets
   
    LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Ticker As String
    Dim Total_Volume As LongLong
    Total_Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim OpeningValue As Double
    OpeningValue = Cells(2, 3).Value
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim ClosingValue As Double
    Dim Greatestpercent As Double
    GratestPercent = Range("K:k").Value
    Dim GreastVolume As Long
    Dim GreatestDec As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    For I = 2 To LastRow
        If Ws.Cells(I + 1, 1).Value <> Ws.Cells(I, 1).Value Then
        
        Ticker = Ws.Cells(I, 1).Value
        Total_Volume = Total_Volume + Ws.Cells(I, 7).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("L" & Summary_Table_Row).Value = Total_Volume
        ClosingValue = Ws.Cells(I, 6).Value
        YearlyChange = (ClosingValue - OpeningValue)
        Range("j" & Summary_Table_Row).Value = YearlyChange
        PercentageChange = YearlyChange / OpeningValue
        Range("k" & Summary_Table_Row).Value = PercentageChange
        Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
        
        Summary_Table_Row = Summary_Table_Row + 1
        OpeningValue = Ws.Cells(I + 1, 3).Value
        Total_Volume = 0
        
        Else: Total_Volume = Total_Volume + Cells(I, 7).Value
        
        If Ws.Cells(I, 10).Value > 0 Then
        Ws.Cells(I, 10).Interior.ColorIndex = 4
        Else
        If Ws.Cells(I, 10).Value < 0 Then
        Ws.Cells(I, 10).Interior.ColorIndex = 3
        Else
        Ws.Cells(I, 10).Interior.ColorIndex = 0
        End If
        End If
         
        
        End If
    Next I
    
Next Ws

End Sub


