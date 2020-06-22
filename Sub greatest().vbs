Sub greatest()

'Presets
Greater_Inc = Cells(2, 11).Value
Greater_Dec = Cells(2, 11).Value
Greatest_Vol = Cells(2, 12).Value

For i = 2 To 2835

    'Greater % Increase conditional
    If Cells(i, 11).Value > Greater_Inc Then
        Greater_Inc = Cells(i, 11).Value
        Greater_Inc_Ticker = Cells(i, 9).Value
    End If
    
    'Greater % Decrease conditional
    If Cells(i, 11).Value < Greater_Dec Then
        Greater_Dec = Cells(i, 11).Value
        Greater_Dec_Ticker = Cells(i, 9).Value
    End If
    
    'Greatest Total Volume conditional
    If Cells(i, 12).Value > Greatest_Vol Then
        Greatest_Vol = Cells(i, 12).Value
        Greatest_Vol_Ticker = Cells(i, 9).Value
    End If

Next i

'Print values with corresponding tickers
Range("O2").Value = Greater_Inc_Ticker
Range("P2").Value = Round(Greater_Inc * 100, 2) & "%"

Range("O3").Value = Greater_Dec_Ticker
Range("P3").Value = Round(Greater_Dec * 100, 2) & "%"

Range("O4").Value = Greatest_Vol_Ticker
Range("P4").Value = Greatest_Vol


End Sub
