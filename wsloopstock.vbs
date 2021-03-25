Attribute VB_Name = "wsloopstock"
Sub wsloopStock()
'Loop for all sheets
Dim ws As Worksheet
    For Each ws In Worksheets

' Label the Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly " + "Change"
        ws.Range("K1").Value = "Percent " + "Change"
        ws.Range("L1").Value = "Total " + "Stock " + "Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
' Labeling the Table
        ws.Range("O2").Value = "Greatest " + "% " + "Increase"
        ws.Range("O3").Value = "Greatest " + "% " + "Decrease"
        ws.Range("O4").Value = "Greatest " + "Total " + "Volume"
' Set a variable for holding the ticker symbol
        Dim tick As String
' Set a variable for holding the total stock volume per ticker symbol
        Dim tsvolume As Double
        tsvolume = 0
'Set variables to calculate yearly change and percent change
        Dim opening As Double
        Dim closing As Double
        Dim ychange As Double
        Dim pchange As Double
'Set variables and initial value for holding Greatest % Increase, % decrease and total volume
        Dim greatinc As Double
        greatinc = 0
        Dim greatdec As Double
        greatdec = 0
        Dim greatvol As Double
        greatvol = 0
'Format % values in table as %s
        ws.Range("Q2", "Q3").NumberFormat = "0.00%"
' define initial opening price
        opening = ws.Range("C2").Value
' keep track of the location for ticker symbol in summary table
        Dim sumtablerow As Integer
        sumtablerow = 2
' get last row
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
' loop through stocks
        For I = 2 To lastrow
' Check if still same ticker symbol
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
' set the ticker symbol
                tick = Cells(I, 1).Value
' add to total stock volume
                tsvolume = tsvolume + ws.Cells(I, 7).Value
' set closing price
                closing = ws.Cells(I, 6).Value
' update ychange
                ychange = closing - opening
' update pchange accounting for dividing by 0
                If opening <> 0 Then
                    pchange = (closing - opening) / opening
' print percent change in summary table
                    ws.Range("K" & sumtablerow).Value = pchange
' format percent change summary column to %
                    ws.Range("K" & sumtablerow).NumberFormat = "0.00%"
                Else
                ws.Range("K" & sumtablerow).Value = "null"
                End If
' print ticker symbol in summary table
                ws.Range("I" & sumtablerow).Value = tick
' print the total stock volume in summary table
                ws.Range("L" & sumtablerow).Value = tsvolume
' print yearly change in summary table
                ws.Range("J" & sumtablerow).Value = ychange
 ' add colour depending of change in value of stock is positive or negative
                If ychange < 0 Then
                    ws.Range("J" & sumtablerow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & sumtablerow).Interior.ColorIndex = 4
                End If
' add one to summary table row
                sumtablerow = sumtablerow + 1
' calculate if new Greatest % increase in stock. If so add to table
                If pchange > greatinc Then
                    greatinc = pchange
                    ws.Range("P2").Value = ws.Cells(I, 1).Value
                    ws.Range("Q2").Value = greatinc
                End If
' calcuate if new Greatest % decrease in stock. If so add to table
                If pchange < greatdec Then
                    greatdec = pchange
                    ws.Range("P3").Value = ws.Cells(I, 1).Value
                    ws.Range("Q3").Value = greatdec
                End If
' calculate if new Total Stock Volume. If so add to table
                If tsvolume > greatvol Then
                    greatvol = tsvolume
                    ws.Range("P4").Value = ws.Cells(I, 1).Value
                    ws.Range("Q4").Value = greatvol
                End If
' reset the total stock volume
                tsvolume = 0
' set next opening price
                opening = ws.Cells(I + 1, 3).Value
' If the cell immediately following is the same ticker symbol then add to ts volume
            Else
                tsvolume = tsvolume + ws.Cells(I, 7).Value
            End If

        Next I

    Next ws

End Sub


