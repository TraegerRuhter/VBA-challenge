Attribute VB_Name = "Module1"
Sub stockstufff()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
  For Each ws In wb.Worksheets

    ' Count rows
    Dim end_of_rows As Double
    end_of_rows = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Variables
    Dim ticker_name As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim yearly_change As Double
    Dim per_change As Double
    Dim total_vol As Double
    Dim year_start As Double
    Dim year_end As Double
    Dim output_row As Double
    Dim greatest_decrease As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_increase As Double
    Dim greatest_increase_ticker As String
    Dim greatest_volume As Double
    Dim greatest_volume_ticker As String
    

    output_row = 2 ' Start output from row 2
    
    
    'set initial vars and the headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    year_start = 2
    total_vol = 0
    ticker_name = ws.Cells(year_start, 1).Value ' Initialize with the first ticker name

    For i = 2 To end_of_rows + 1
        If ws.Cells(i, 1).Value <> ticker_name Then
            year_end = i - 1

            ' Get opening and closing prices
            Open_Price = ws.Cells(year_start, 3).Value
            Close_Price = ws.Cells(year_end, 6).Value

            ' Calculate yearly_change and per_change
            yearly_change = Close_Price - Open_Price
            If Open_Price <> 0 Then
                per_change = (yearly_change / Open_Price) * 100
            Else
                per_change = 0
            End If

            ' Output results on the next row
            ws.Cells(output_row, 9).Value = ticker_name
            ws.Cells(output_row, 10).Value = yearly_change
            
            'color cells
                If yearly_change < 0 Then
                ws.Cells(output_row, 10).Interior.ColorIndex = 3 ' Red
                ElseIf yearly_change > 0 Then
                 ws.Cells(output_row, 10).Interior.ColorIndex = 4 ' Green
                 Else
                    ws.Cells(output_row, 10).Interior.ColorIndex = 6 ' Yellow
                    End If
                    
            'keep track of the biggest and lowest change
            
            If per_change > greatest_increase Then
                greatest_increase = per_change
                greatest_increase_ticker = ticker_name
            ElseIf per_change < greatest_decrease Then
                greatest_decrease = per_change
                greatest_decrease_ticker = ticker_name
            End If

            If total_vol > greatest_volume Then
                greatest_volume = total_vol
                greatest_volume_ticker = ticker_name
            End If
            
            ws.Cells(output_row, 11).Value = per_change
            ws.Cells(output_row, 12).Value = total_vol

            'next row
            output_row = output_row + 1

            ' Reset total_vol and year_start
            total_vol = 0
            year_start = i

            ' Update the ticker_name
            ticker_name = ws.Cells(i, 1).Value
        End If

        ' total volume for the current year
        total_vol = total_vol + ws.Cells(i, 7).Value
    Next i

    'output greates increase, decrease, and volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ws.Cells(2, 16).Value = greatest_increase_ticker
    ws.Cells(3, 16).Value = greatest_decrease_ticker
    ws.Cells(4, 16).Value = greatest_volume_ticker
    
    ws.Cells(2, 17).Value = greatest_increase
    ws.Cells(3, 17).Value = greatest_decrease
    ws.Cells(4, 17).Value = greatest_volume


Next ws


End Sub
