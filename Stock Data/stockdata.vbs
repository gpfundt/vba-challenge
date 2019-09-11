Sub stockData()

'If ActiveSheet.Index = Worksheets.Count Then
'    Worksheets(1).Activate
'Else
'    ActiveSheet.Next.Activate
'End If

'Define the variables
Dim tickercoulmn As String
Dim counter As Long
counter = 2
Dim openAmt As Double
Dim closingAmt As Double
Dim yearlyChange As Double
Dim volumn As Double
Dim percentChange As Double
Dim realPercent As String
Dim wsYear As Double
Dim openDate As Double

'Make the headers for the new columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'find last row
lastRow = Cells.SpecialCells(xlCellTypeLastCell).Row

'find year
wsYear = Left(Range("B2"), 4)
'find open date based on the year defined on the worksheet
openDate = wsYear & Right("0000" & 101, 4)
    
    For i = 2 To lastRow
        'Find the opening amount at the beginning of the year
        If Cells(i, 2).Value = openDate Then
            openAmt = Cells(i, 3).Value
        End If
        'If openAmt is 0 the true opening amount is the first amount of the year that is not zero
        If openAmt = 0 Then
            openAmt = openAmt + Cells(i + 1, 3).Value
        End If
        
        'Find the tickers that are different
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            
            'Make ticker column for every different ticker
            tickercolumn = Cells(i, 1).Value
            Cells(counter, 9).Value = tickercolumn
            
            'Find closing amount at the end of the year
            closingAmt = Cells(i, 6).Value

            'Initialize the volum for each ticker
            volumn = volumn + Cells(i, 7)
            Cells(counter, 12).Value = volumn
            
            'Calculate yearly change
            yearlyChange = closingAmt - openAmt
            Cells(counter, 10).Value = yearlyChange
            
            'Calculate the percent change for each ticker
            percentChange = yearlyChange / openAmt
            'Format percent change column as %
            realPercent = FormatPercent(percentChange, 0)
            Cells(counter, 11).Value = realPercent
            
            'Reset volumn
            volumn = 0
        
            'color the positive changes green and the negative changes in red
            If Cells(counter, 10).Value > 0 Then
                Cells(counter, 10).Interior.Color = RGB(0, 225, 0)
            Else
                Cells(counter, 10).Interior.Color = RGB(225, 0, 0)
            End If
                        
            'Increment counter
            counter = counter + 1
            
        Else 'To add up the volumn when the tickers are the same
            
            'Add total volumn for each ticker
            Cells(counter, 12).Value = volumn
            volumn = volumn + Cells(i, 7)
        End If
    Next i

Dim max1 As Double
Dim min1 As Double
Dim maxV As Double
    
Range("N2").Value = "Greatest % Increase"
max1 = Application.WorksheetFunction.Max(Columns("k"))
Range("O2").Value = FormatPercent(max1, 0)

Range("N3").Value = "Greatest % Decrease"
min1 = Application.WorksheetFunction.Min(Columns("n"))
Range("O3").Value = FormatPercent(min1, 0)

Range("N4").Value = "Greatest Volumn"
maxV = Application.WorksheetFunction.Max(Columns("l"))
Range("O4").Value = maxV

End Sub
