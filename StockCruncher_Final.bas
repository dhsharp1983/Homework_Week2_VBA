Sub StockCruncher()

    Dim DateString As String
    Dim DateYear As Integer
    Dim DateMonth As Integer
    Dim DateDay As Integer
    Dim UniqueTickerCounter As Integer
    Dim TotalStockVolume As LongLong
    Dim TickerStartingRow As Long
    Dim TickerEndingRow As Long
    Dim OpeningStockPrice As Double
    Dim ClosingStockPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalRows As Long
    Dim WorkSheetCount As Integer
    Dim ws As Worksheet
    
    
    
    
    'start worksheet loop
    For Each ws In Worksheets
        
        ws.Activate
    
        'Define total number of rows in dataset
        TotalRows = ws.Range("A1").End(xlDown).Row
        
        'Set certain counters
        'UniqueTickerCounter is for Display Table to start at row2
        UniqueTickerCounter = 1
        TotalStockVolume = 0
        TickerStartingRow = 0

        'Manually add column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Columns("I:P").EntireColumn.AutoFit
        
        'Start Data Crunch Loop
        For i = 2 To TotalRows
        
            'Store First Row as TickerStartingRow
            If TickerStartingRow = 0 Then
                TickerStartingRow = i
            End If
            
            'Increment TotalStockVolume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            'If Next Rows Ticker is Different, perform stock calculations
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                
                'Set TickerEndingRow
                TickerEndingRow = i
                
                'Increment UniqueTickerCounter for Display Table
                UniqueTickerCounter = UniqueTickerCounter + 1
                
                'Calculate Open / Close Stock Prices
                OpeningStockPrice = ws.Cells(TickerStartingRow, 3).Value
                ClosingStockPrice = ws.Cells(TickerEndingRow, 6).Value
                
                'Calculate Stock Price Changes
                YearlyChange = ClosingStockPrice - OpeningStockPrice
                'Calculate Percent Change, needs error handling to avoid Div0
                If ClosingStockPrice <> 0 And OpeningStockPrice <> 0 Then
                    PercentChange = (ClosingStockPrice - OpeningStockPrice) / OpeningStockPrice
                    End If
                
                'Start Outputting Values to Display Table
                ws.Cells(UniqueTickerCounter, 9).Value = ws.Cells(i, 1)
                ws.Cells(UniqueTickerCounter, 10).Value = YearlyChange
                ws.Cells(UniqueTickerCounter, 11).Value = PercentChange
                ws.Cells(UniqueTickerCounter, 12).Value = TotalStockVolume
                
                'Format Rows of Display Table
                ws.Cells(UniqueTickerCounter, 10).NumberFormat = "$#,##0.00"
                ws.Cells(UniqueTickerCounter, 11).NumberFormat = "0.00%"
                
                'Apply Conditional Green/Red Format ws.Cells
                If ws.Cells(UniqueTickerCounter, 10).Value < 0 Then
                    ws.Cells(UniqueTickerCounter, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(UniqueTickerCounter, 10).Value > 0 Then
                    ws.Cells(UniqueTickerCounter, 10).Interior.ColorIndex = 4
                    
                    End If
                
                'Reset certain variables
                YearlyChange = 0
                PercentChange = 0
                TickerStartingRow = 0
                TotalStockVolume = 0
                TickerEndingRow = 0
                
                'End "If next ticker row is different"
                End If
            
            'End Data Parsing Loop
            Next
            
            'Start Greatest Figure Analysis
                'Definte Sumary Rows Variable
                Dim TotalSummaryRows As Long
                
                'Calculate number of rows in Summary Display Table
                TotalSummaryRows = ws.Range("J1").End(xlDown).Row
                
                'Define Arrays
                Dim PercentChangeArray() As Variant
                Dim TotalStockVolumesArray() As Variant
                Dim TickerArray() As Variant
                
                'ReDim Arrays
                ReDim PercentChangeArray(TotalSummaryRows)
                ReDim TotalStockVolumesArray(TotalSummaryRows)
                ReDim TickerArray(TotalSummaryRows)
                
                'Read Data Into Arrays
                For ii = 2 To TotalSummaryRows
                    PercentChangeArray(ii) = ws.Cells(ii, 11).Value
                    TotalStockVolumesArray(ii) = ws.Cells(ii, 12).Value
                    TickerArray(ii) = ws.Cells(ii, 9).Value
                    Next
                    
                'Apply Formulas to Array Data
                ws.Range("P2").Value = WorksheetFunction.Max(PercentChangeArray)
                ws.Range("P3").Value = WorksheetFunction.Min(PercentChangeArray)
                ws.Range("P4").Value = WorksheetFunction.Max(TotalStockVolumesArray)
                
                'Find Associated Tickers to Values
                'First match the Greatest % Increase
                For iii = 2 To TotalSummaryRows
                    If ws.Cells(iii, 11).Value = ws.Range("P2").Value Then
                        ws.Range("O2").Value = ws.Cells(iii, 9).Value
                    End If
                    Next
                    
                    
                'Next match the Greatest % Decrease
                For iii = 2 To TotalSummaryRows
                    If ws.Cells(iii, 11).Value = ws.Range("P3").Value Then
                        ws.Range("O3").Value = ws.Cells(iii, 9).Value
                    End If
                    Next
                    
                'Next match the Total Stock Volumes
                For iii = 2 To TotalSummaryRows
                    If ws.Cells(iii, 12).Value = ws.Range("P4").Value Then
                        ws.Range("O4").Value = ws.Cells(iii, 9).Value
                    End If
                    Next
                    
                'Format Percentage Numbers
                ws.Range("P2:P3").NumberFormat = "0.00%"
        
        
    Next ws

End Sub



