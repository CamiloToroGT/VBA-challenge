Attribute VB_Name = "Module1"
Sub stockAnalysis()
 Dim ws As Worksheet
    Dim year_opening As Single
    Dim year_closing As Single
    Dim volume As Double
    Dim select_index As Double
    Dim first_row As Double
    Dim select_row As Double
    Dim last_row As Double

    
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        volume = 0
        select_index = 2
        first_row = 2
        select_row = 2
        last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        
        
        'step 1: Put headings/labels to columns and rows
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'step 2: Loop through all rows to find unique tickers and place each unique ticker to column 9 (Tickers)
        
        For i = first_row To last_row
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers <> tickers2 Then
                Cells(select_row, 9).Value = tickers
                select_row = select_row + 1
            End If
         Next i
    
        'step 3: Loop through all rows while adding the volume if the ticker is the same as the previous ticker.
        'Once ticker has changed, reset volume to 0 and continue.
        
        For i = first_row To last_row + 1
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers = tickers2 And i > 2 Then
                volume = volume + Cells(i, 7).Value
            ElseIf i > 2 Then
                Cells(select_index, 12).Value = volume
                select_index = select_index + 1
                volume = 0
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
            
        'Step 4: Loop through all rows and check the previous ticker and the next ticker.
        'If the previous ticker is different then assign year_opening.
        'If the next ticker is different, assign year_closing.
        
        select_index = 2
        For i = first_row To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                year_closing = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                year_opening = Cells(i, 3).Value
            End If
            If year_opening > 0 And year_closing > 0 Then
                increase = year_closing - year_opening
                percent_increase = increase / year_opening
                Cells(select_index, 10).Value = increase
                Cells(select_index, 11).Value = FormatPercent(percent_increase)
                year_closing = 0
                year_opening = 0
                select_index = select_index + 1
            End If
        Next i
        
        'Loops through percent change column then highlights either green for a positive change or red for a negative change
        For i = first_row To last_row
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        'Find minimum and maximum values in the percent change column.
        max_percentage = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        min_percentage = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        max_volume = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        'Assigns each value from above to their corresponding cells under the value column
        
        Range("Q2").Value = FormatPercent(max_percentage)
        Range("Q3").Value = FormatPercent(min_percentage)
        Range("Q4").Value = max_volume
        
        
        'Loops through the Percent Change & the Total stock volume column .
        'If either column contains minimum or maximum values, apply corresponding ticker to corresponding cell in column 16 under ticker.
        
        For i = first_row To last_row
            If max_percentage = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf min_percentage = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            ElseIf max_volume = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        
    Next ws
End Sub


