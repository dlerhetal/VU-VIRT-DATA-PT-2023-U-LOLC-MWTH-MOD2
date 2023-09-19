Attribute VB_Name = "Module1"
Sub StockSum()
    
    'Place ws. in front of Cell() or Range() to apply these commands to the entire workbook.
    For Each ws In Worksheets
        
        'Define what is the last row for each new worksheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'ThisRow outputs the Ticker, YearlyChange, PercentChange, and TotalStockVolume.
            Dim ThisRow As Double
                ThisRow = 2
            Dim StarterRow As Double
                StarterRow = 2
            Dim ValueRow As Double
                ValueRow = 2
            Dim DailyVolume As Double
            Dim TotalVolume As Double
                TotalVolume = 0
            Dim VolumeRow As Double
                VolumeRow = 0
            Dim MinValueRow As Double
                MinValueRow = 0
            Dim MaxValueRow As Double
                MaxValueRow = 0
            Dim VolumeRange As Range
            Set VolumeRange = ws.Range("L1:L" & LastRow)
            Dim YearlyChangeRange As Range
            Set YearlyChangeRange = ws.Range("J1:J" & LastRow)
            Dim MaxYearlyChangeRange As Range
            Set MaxYearlyChangeRange = ws.Range("Q2")
            Dim MinYearlyChangeRange As Range
            Set MinYearlyChangeRange = ws.Range("Q3")
            Dim MaxVolumeRange As Range
            Set MaxVolumeRange = ws.Range("Q4")
            
            'Print the column headings and 3 row titles.
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Volume"
            ws.Range("O2").Value = "Greatest%Increase"
            ws.Range("O3").Value = "Greatest%Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("Q2").Value = MaxPercent
            ws.Range("Q3").Value = MinPercent
            ws.Range("Q4").Value = MaxVolume
            
            
    'Loop through each row.
    For NextRow = 2 To LastRow
                
        'Compare TickerName with NextTickerName to find unique values in column A.
        TickerName = ws.Cells(NextRow, 1).Value
        NextTickerName = ws.Cells(NextRow + 1, 1).Value
        DailyVolume = ws.Cells(NextRow, 7).Value
                        
        'When a value in column A does not match the value in the next row in column A...
        If NextTickerName <> TickerName Then
                        
            'Print the last unique TickerName value in column A to column I.
            ws.Cells(ThisRow, 9) = ws.Cells(NextRow, 1).Value
                        
            'Add the final row in column G that corresponds with the last unique TickerName value in column A.
            TotalVolume = TotalVolume + DailyVolume

            'Set the value of column J
            ws.Cells(ThisRow, 10).Value = ws.Cells(NextRow, 6).Value - ws.Cells(StarterRow, 3).Value

            'Set the value of percent change
            ws.Cells(ThisRow, 11).Value = (ws.Cells(ThisRow, 10).Value / ws.Cells(StarterRow, 3).Value) * 100

            'Now that we have printed all of the information to the output columns, go to the next row.
            ws.Cells(ThisRow, 12).Value = TotalVolume
            TotalVolume = 0
            ThisRow = ThisRow + 1
            StarterRow = NextRow + 1

            'While we are iterating through column A, looking for the next unique value, continue adding up the stock volume.
            Else
            TotalVolume = TotalVolume + DailyVolume
                
        End If
                                        
        Next NextRow
            
    Next ws
    
        MsgBox "Updated!"
    
End Sub
            
Sub ColorHandler()

    For Each ws In Worksheets
        
        'Define what is the last row for each new worksheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'ThisRow outputs the Ticker, YearlyChange, PercentChange, and TotalStockVolume.
            Dim ThisRow As Double
                ThisRow = 2
            Dim StarterRow As Double
                StarterRow = 2
            Dim ValueRow As Double
                ValueRow = 2
            Dim DailyVolume As Double
            Dim TotalVolume As Double
                TotalVolume = 0
            Dim VolumeRow As Double
                VolumeRow = 0
            Dim MinValueRow As Double
                MinValueRow = 0
            Dim MaxValueRow As Double
                MaxValueRow = 0
            Dim VolumeRange As Range
            Set VolumeRange = ws.Range("L1:L" & LastRow)
            Dim YearlyChangeRange As Range
            Set YearlyChangeRange = ws.Range("J1:J" & LastRow)
            Dim MaxYearlyChangeRange As Range
            Set MaxYearlyChangeRange = ws.Range("Q2")
            Dim MinYearlyChangeRange As Range
            Set MinYearlyChangeRange = ws.Range("Q3")
            Dim MaxVolumeRange As Range
            Set MaxVolumeRange = ws.Range("Q4")
            
            'Print the column headings and 3 row titles.
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Volume"
            ws.Range("O2").Value = "Greatest%Increase"
            ws.Range("O3").Value = "Greatest%Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("Q2").Value = MaxPercent
            ws.Range("Q3").Value = MinPercent
            ws.Range("Q4").Value = MaxVolume

        'This iteration is for conditional formatting for column J.
        For ColorRow = 2 To LastRow
                    
            'If the cell in column J is blank, do nothing.
            If ws.Range("J" & ColorRow).Value = 0 Then
                ws.Cells(ColorRow, 10).Interior.ColorIndex = 0
                    
            'If the cell in column J is positive, set the fill color to green.
            ElseIf ws.Range("J" & ColorRow).Value > 0 Then
                ws.Cells(ColorRow, 10).Interior.ColorIndex = 4
            
            'If the cell in column J is negative, set the fill color to red.
            ElseIf ws.Range("J" & ColorRow).Value < 0 Then
                ws.Cells(ColorRow, 10).Interior.ColorIndex = 3
                        
            Else
           
            End If

        Next ColorRow

    Next ws
    
        MsgBox "Updated!"
    
End Sub
Sub MaxValueHandler()

    'Place ws. in front of Cell() or Range() to apply these commands to the entire workbook.
    For Each ws In Worksheets
        
        'Define what is the last row for each new worksheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'ThisRow outputs the Ticker, YearlyChange, PercentChange, and TotalStockVolume.
            Dim ThisRow As Double
                ThisRow = 2
            Dim StarterRow As Double
                StarterRow = 2
            Dim ValueRow As Double
                ValueRow = 2
            Dim DailyVolume As Double
            Dim TotalVolume As Double
                TotalVolume = 0
            Dim VolumeRow As Double
                VolumeRow = 0
            Dim MinValueRow As Double
                MinValueRow = 0
            Dim MaxValueRow As Double
                MaxValueRow = 0
            Dim VolumeRange As Range
            Set VolumeRange = ws.Range("L1:L" & LastRow)
            Dim YearlyChangeRange As Range
            Set YearlyChangeRange = ws.Range("J1:J" & LastRow)
            Dim MaxYearlyChangeRange As Range
            Set MaxYearlyChangeRange = ws.Range("Q2")
            Dim MinYearlyChangeRange As Range
            Set MinYearlyChangeRange = ws.Range("Q3")
            Dim MaxVolumeRange As Range
            Set MaxVolumeRange = ws.Range("Q4")


        'This iteration discovers and prints the Greatest%Increase from column J.
        For MaxValueRow = 2 To LastRow
        MaxYearlyChangeRange.Value = Application.WorksheetFunction.Max(YearlyChangeRange)

            If ws.Cells(MaxValueRow, 10).Value = MaxYearlyChangeRange.Value Then
                ws.Range("P2").Value = ws.Cells(MaxValueRow, 9).Value
        
        Else
        
        End If
        
        Next MaxValueRow
            
    Next ws
            
End Sub

Sub MinValueHandler()

    'Place ws. in front of Cell() or Range() to apply these commands to the entire workbook.
    For Each ws In Worksheets
        
        'Define what is the last row for each new worksheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'ThisRow outputs the Ticker, YearlyChange, PercentChange, and TotalStockVolume.
            Dim ThisRow As Double
                ThisRow = 2
            Dim StarterRow As Double
                StarterRow = 2
            Dim ValueRow As Double
                ValueRow = 2
            Dim DailyVolume As Double
            Dim TotalVolume As Double
                TotalVolume = 0
            Dim VolumeRow As Double
                VolumeRow = 0
            Dim MinValueRow As Double
                MinValueRow = 0
            Dim MaxValueRow As Double
                MaxValueRow = 0
            Dim VolumeRange As Range
            Set VolumeRange = ws.Range("L1:L" & LastRow)
            Dim YearlyChangeRange As Range
            Set YearlyChangeRange = ws.Range("J1:J" & LastRow)
            Dim MaxYearlyChangeRange As Range
            Set MaxYearlyChangeRange = ws.Range("Q2")
            Dim MinYearlyChangeRange As Range
            Set MinYearlyChangeRange = ws.Range("Q3")
            Dim MaxVolumeRange As Range
            Set MaxVolumeRange = ws.Range("Q4")
        
        'This iteration discovers and prints the Greatest%Decrease from column J.
        For MinValueRow = 2 To LastRow
        MinYearlyChangeRange.Value = Application.WorksheetFunction.Min(YearlyChangeRange)
        
            If ws.Cells(MinValueRow, 10).Value = MinYearlyChangeRange.Value Then
                ws.Range("P3").Value = ws.Cells(MinValueRow, 9).Value
        
        Else
        
        End If
        
        Next MinValueRow
        
        MsgBox "Updated!"
        
    Next ws

End Sub

Sub VolumeHandler()

    'Place ws. in front of Cell() or Range() to apply these commands to the entire workbook.
    For Each ws In Worksheets
        
        'Define what is the last row for each new worksheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'ThisRow outputs the Ticker, YearlyChange, PercentChange, and TotalStockVolume.
            Dim ThisRow As Double
                ThisRow = 2
            Dim StarterRow As Double
                StarterRow = 2
            Dim ValueRow As Double
                ValueRow = 2
            Dim DailyVolume As Double
            Dim TotalVolume As Double
                TotalVolume = 0
            Dim VolumeRow As Long
                VolumeRow = 0
            Dim MinValueRow As Double
                MinValueRow = 0
            Dim MaxValueRow As Double
                MaxValueRow = 0
            Dim VolumeRange As Range
            Set VolumeRange = ws.Range("L1:L" & LastRow)
            Dim YearlyChangeRange As Range
            Set YearlyChangeRange = ws.Range("J1:J" & LastRow)
            Dim MaxYearlyChangeRange As Range
            Set MaxYearlyChangeRange = ws.Range("Q2")
            Dim MinYearlyChangeRange As Range
            Set MinYearlyChangeRange = ws.Range("Q3")
            Dim MaxVolumeRange As Range
            Set MaxVolumeRange = ws.Range("Q4")

        'This iteration discovers and prints the Greatest Total Volume from column L.
        For VolumeRow = 2 To LastRow
        MaxVolumeRange.Value = Application.WorksheetFunction.Max(VolumeRange)
        
            If ws.Cells(VolumeRow, 12).Value = MaxVolumeRange.Value Then
                ws.Range("P4").Value = ws.Cells(VolumeRow, 9).Value
        
        Else
        
        End If

        Next VolumeRow
                    
    Next ws
    
        MsgBox "Updated!"
    
End Sub

Sub TestReset()

For Each ws In Worksheets

    ws.Range("I1:AB22271").ClearContents
    ws.Range("J1:J222271").Interior.ColorIndex = 0

Next ws
    MsgBox "Clear!"

End Sub



