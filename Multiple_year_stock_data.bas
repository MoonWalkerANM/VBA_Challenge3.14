Attribute VB_Name = "Module1"
Option Explicit
Sub Stock_Data()

        Dim ws As Worksheet
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim TickerPos As Long
        Dim i As Long
        Dim tickerCell As Range
        Dim tickerRange As Range
        Dim startRow As Long
        Dim currentRow As Long
        Dim lastRow As Long
        Dim rowData As Range
        Dim DataRange As Range
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim Value As Double

        
        
        For Each ws In ThisWorkbook.Sheets
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            Ticker = ""
            TotalStockVolume = 0
            startRow = 2
            currentRow = 2
            TickerPos = 2
            Set tickerRange = ws.Range("A1:A753001")
            'Add new columns
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "YearlyChange"
            ws.Cells(1, 11).Value = "PercentChange"
            ws.Cells(1, 12).Value = "TotalStockVolume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            

            For i = 2 To lastRow
                If Ticker = "" Then
                    Ticker = ws.Cells(i, 1).Value
                    startRow = ws.Cells(i, 1).Row
                End If
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                currentRow = currentRow + 1
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    startRow = ws.Cells(i - 250, 1).Row
                    ws.Cells(TickerPos, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(TickerPos, 10).Value = ws.Cells(i, 6).Value - ws.Cells(startRow, 3).Value
                    If ws.Cells(TickerPos, 10).Value > 0 Then
                        ws.Cells(TickerPos, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(TickerPos, 10).Interior.ColorIndex = 3
                    End If
                    ws.Cells(TickerPos, 11).Value = (ws.Cells(TickerPos, 10).Value / ws.Cells(startRow, 3).Value)
                    ws.Cells(TickerPos, 11).NumberFormat = "0.00%"
                    ws.Cells(TickerPos, 12).Value = TotalStockVolume
                    Ticker = ws.Cells(i + 1, 1).Value
                    TotalStockVolume = 0
                    TickerPos = TickerPos + 1
                End If
            
              Next i
           
        Next ws
        
End Sub
