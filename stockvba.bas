Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        SummaryRow = 2
        
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
        MaxIncreaseTicker = ""
        MaxDecreaseTicker = ""
        MaxVolumeTicker = ""
                ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
            Ticker = ws.Cells(i, 1).Value
            ClosePrice = ws.Cells(i, 6).Value
            
            If ws.Cells(i - 1, 1).Value <> Ticker Then
                OpenPrice = ws.Cells(i, 3).Value
            End If
            
            YearlyChange = ClosePrice - OpenPrice
            If OpenPrice <> 0 Then
                PercentChange = (YearlyChange / OpenPrice) * 100
            Else
                PercentChange = 0
            End If
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = PercentChange
            ws.Cells(SummaryRow, 12).Value = TotalVolume
            
            If YearlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) 
            ElseIf YearlyChange < 0 Then
                ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) 
            End If
            
            If PercentChange > MaxIncrease Then
                MaxIncrease = PercentChange
                MaxIncreaseTicker = Ticker
            ElseIf PercentChange < MaxDecrease Then
                MaxDecrease = PercentChange
                MaxDecreaseTicker = Ticker
            End If
            If TotalVolume > MaxVolume Then
                MaxVolume = TotalVolume
                MaxVolumeTicker = Ticker
            End If
            
            SummaryRow = SummaryRow + 1
        Next i
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(2, 17).Value = MaxIncrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = MaxDecreaseTicker
        ws.Cells(3, 17).Value = MaxDecrease & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(4, 17).Value = MaxVolume
    Next ws
End Sub
