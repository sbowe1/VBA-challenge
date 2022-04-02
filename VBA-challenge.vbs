Sub Stocks()
    ' A for loop to scroll through each sheet in the file
    Dim ws as Worksheet
    For Each ws In Worksheets
        ' Creating new column headers on the sorted sheet
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Columns("J:L").AutoFit
            ' Adding new column/row headers for 'greatest' categories
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            .Columns("O:Q").AutoFit
        End With

        ' Last row of sheet after each paste (+1 to get the first empty row)
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1

        ' Finding each unique ticker from column A
        ' Last row of the column A of ws
        last = ws.Cells(Rows.Count, "A").End(xlUp).Row
        ' Copying ticker range from ws to column I of the sorted worksheet
        ws.Range("A2" & ":A" & last).Copy ws.Range("I" & last_row)
        ws.Range("I:I").RemoveDuplicates Columns:=1

        last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        max_increase = 0
        max_decrease = 0
        max_volume = 0
         ' For loop that goes through each ticker in column I
        For i = 2 To last_row2
            ' Idenitifying the rows corresponding to ticker i in column A of ws 
            open_row = ws.Range("A:A").Find(what:=ws.Cells(i, 9).Value, after:=ws.Range("A1")).Row
            close_row = ws.Range("A:A").Find(what:=ws.Cells(i, 9).Value, after:=ws.Range("A1"), searchdirection:=xlPrevious).Row
            ' Yearly change is final close - initial open
            yearly_change = ws.Cells(close_row, 6).Value - ws.Cells(open_row, 3).Value
            ws.Cells(i, 10) = yearly_change
            ' Conditional formatting highlight positive change green and negative change red
            If ws.Cells(i, 10).Value > 0 Then 
                ' Assigning the color green to positive changes
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ' Assigning the color red to negative changes
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If

            ' Percent change is (final - open)/open, or yearly/open
            percent_change = yearly_change/ws.Cells(open_row, 3).Value
            ws.Cells(i, 11).Value = percent_change
            ' Making the value appear as a percentage 
            ws.Cells(i, 11).NumberFormat = "0.00%"

            ' Total ticker volume as the sum of column G for each ticker i 
            Dim total_volume As Variant
            total_volume = WorksheetFunction.Sum(ws.Range("G" & open_row & ":G" & close_row))
            ws.Cells(i, 12).Value = total_volume

            ' Identifying the 'greatest' values for the bonus section
            If ws.Cells(i, 11) > max_increase Then
                ' Make the new max_increase equal to the value of Cell(i, 11)
                max_increase = ws.Cells(i, 11)
                ' Obtain according ticker value
                increase_ticker = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11) < max_decrease Then
                ' Make the new max_decrease equal to the value of Cell(i, 11)
                max_decrease = ws.Cells(i, 11)
                ' Obtain according ticker value
                decrease_ticker = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value > max_volume Then
                ' Make the new max_volume equal to the value of Cell(i, 12)
                max_volume = ws.Cells(i, 12).Value
                ' Obtain according ticker value
                volume_ticker = ws.Cells(i, 9).Value
            End If
        Next i

        ' Putting each 'greatest' value into its assigned cell
        ws.Range("P2").Value = increase_ticker
        ws.Range("P3").Value = decrease_ticker
        ws.Range("P4").Value = volume_ticker
        ws.Range("Q2").Value = max_increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = max_decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = max_volume
    Next ws
End Sub    
