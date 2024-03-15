Attribute VB_Name = "Module1"
Option Explicit
Sub Stock_Data()

        Dim ws As Worksheets
        Dim WorksheetName As String
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim TickerPos As Integer
        Dim i As Long
        Dim tickercell As Range
        Dim startRow As Long
        Dim tickerRange As Range
        
'        Dim lastrow As Long
'        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        
        TickerPos = 2
        Ticker = ""
        TotalStockVolume = 0

        
        'Add new columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "YearlyChange"
        Cells(1, 11).Value = "PercentChange"
        Cells(1, 12).Value = "TotalStockVolume"


    For i = 2 To 753001
        If Ticker = "" Then
        Ticker = Cells(i, 1).Value
        End If
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        Set tickerRange = Range("A1:A753001")
        For Each tickercell In tickerRange
        startRow = tickercell.Row
    
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Cells(TickerPos, 9).Value = Cells(i, 1).Value
                Cells(TickerPos, 12).Value = TotalStockVolume
                Cells(TickerPos, 10).Value = Cells(tickercell.Row, 3).Value - Cells(i, 6).Value
                'Cells(TickerPos, 10).Value = Val(Cells(currentrow - 1, 6).Value) - Val(ws.Cells(currentrow, 3).Value)
                Cells(TickerPos, 11).Value = Cells(i, 6).Value / Cells(i, 3).Value
                Ticker = Cells(i + 1, 1).Value
                TotalStockVolume = 0
                TickerPos = TickerPos + 1
            End If
        Next tickercell
                    
    Next i
        
End Sub
