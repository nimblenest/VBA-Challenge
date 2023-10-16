VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateStockInfo()
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
    
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestvolume As Double
        Dim tickerGreatestIncrease As String
        Dim tickerGreatestDecrease As String
        Dim tickerGreatestVolume As String
        
        greatestIncrease = -10
        greatestDecrease = 999999
        greatestvolume = 0
    
        Dim i As Long
        For i = 2 To lastRow
           
            Dim ticker As String
            Dim openingPrice As Double
            Dim closingPrice As Double
            Dim volume As Double
    
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            closingPrice = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value
             
            Dim yearlyChange As Double
            Dim percentageChange As Double
            yearlyChange = closingPrice - openingPrice
            percentageChange = ((closingPrice - openingPrice) / openingPrice) * 100
            
            ws.Cells(i, 9).Value = ticker
            ws.Cells(i, 10).Value = yearlyChange
            ws.Cells(i, 11).Value = percentageChange
            ws.Cells(i, 12).Value = volume
            
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
                
           ws.Cells(i, 11).NumberFormat = "0.00%"
           
           If percentageChange > greatestIncrease Then
               greatestIncrease = percentageChange
               tickerGreatestIncrease = ticker
           End If
           
           If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                tickerGreatestDecrease = ticker
           End If
           
           If volume > greatestvolume Then
                greatestvolume = volume
                tickerGreatestVolume = ticker
           End If
           
        Next i
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = " Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = greatestvolume
        ws.Cells(2, 16).Value = tickerGreatestIncrease
        ws.Cells(3, 16).Value = tickerGreatestDecrease
        ws.Cells(4, 16).Value = tickerGreatestVolume
   Next ws
End Sub
