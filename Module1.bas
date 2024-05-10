Attribute VB_Name = "Module1"
Sub CalculateAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Call CalculateQuarterlyStockData(ws)
    Next ws
End Sub

Sub CalculateQuarterlyStockData(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim ticker As String
    Dim startPrice As Double
    Dim endPrice As Double
    Dim totalVolume As Double
    Dim quarterChange As Double
    Dim percentChange As Double
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxVolume As Double
    Dim tickerMaxIncrease As String
    Dim tickerMaxDecrease As String
    Dim tickerMaxVolume As String
    Dim startRow As Integer
    Dim i As Long
    
    startRow = 2
    ticker = ws.Cells(startRow, 1).Value
    startPrice = ws.Cells(startRow, 3).Value
    totalVolume = 0
    maxPercentIncrease = -100
    maxPercentDecrease = 100
    maxVolume = 0

    ' Output headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarter Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

    Dim outputRow As Integer
    outputRow = 2

    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
            endPrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            quarterChange = endPrice - startPrice
            If startPrice <> 0 Then
                percentChange = (quarterChange / startPrice) * 100
            Else
                percentChange = 0
            End If

            ' Update max/min values
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                tickerMaxIncrease = ticker
            End If
            If percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                tickerMaxDecrease = ticker
            End If
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                tickerMaxVolume = ticker
            End If

            ' Output the results
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterChange
            ws.Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
            ws.Cells(outputRow, 12).Value = totalVolume

            ' Prepare for the next ticker
            outputRow = outputRow + 1
            If i + 1 <= lastRow Then
                ticker = ws.Cells(i + 1, 1).Value
                startPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            End If
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i

    ' Output greatest values
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = tickerMaxIncrease
    ws.Cells(2, 16).Value = Format(maxPercentIncrease, "0.00") & "%"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = tickerMaxDecrease
    ws.Cells(3, 16).Value = Format(maxPercentDecrease, "0.00") & "%"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Value = tickerMaxVolume
    ws.Cells(4, 16).Value = maxVolume

    ' Add Conditional Formatting for "Quarter Change"
    Dim quarterChangeRange As Range
    Set quarterChangeRange = ws.Range(ws.Cells(2, 10), ws.Cells(outputRow - 1, 10))
    quarterChangeRange.FormatConditions.Delete

    Dim condNeg As FormatCondition
    Set condNeg = quarterChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    condNeg.Font.Color = RGB(255, 0, 0)
    condNeg.Interior.Color = RGB(255, 204, 204)

    Dim condPos As FormatCondition
    Set condPos = quarterChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
    condPos.Font.Color = RGB(0, 128, 0)
    condPos.Interior.Color = RGB(204, 255, 204)
End Sub

Sub Clear()
    ' This clears the contents of cells from A1 to C10 on the active worksheet.
    Worksheets("Q1").Range("New_data1").Clear
    Worksheets("Q2").Range("New_data2").Clear
    Worksheets("Q3").Range("New_data3").Clear
    Worksheets("Q4").Range("New_data4").Clear
End Sub

