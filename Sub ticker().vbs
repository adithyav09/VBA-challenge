Sub ticker()
Dim i As Long
Dim varray As Variant

'varray helps me loop through all rows without runtime error (found online).
varray = Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row).Value

'Preset variables: (j represents individual tickers row in ticker column)
j = 2
Yearly_Change = 0
Open_Value = Cells(2, 3).Value
T_Volume = 0

For i = 2 To UBound(varray, 1)

'Total Stock Volume Iterator
T_Volume = Cells(i, 7).Value + T_Volume
    
    'This condition checks when the next ticker is and computes all the desired info
    'for that ticker's column.
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'Ticker Column
        Cells(j, 9).Value = Cells(i, 1).Value
        
        'Yearly Change
        Yearly_Change = Cells(i, 6).Value - Open_Value
        Cells(j, 10).Value = Yearly_Change
        
        'Yearly Change Cell Color
        If Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        Else
            Cells(j, 10).Interior.ColorIndex = 4
        End If
        
        'Percent Change
        Percent_Change = ((Cells(i, 6).Value / Open_Value) - 1)
        Cells(j, 11).Value = Round(Percent_Change * 100, 2) & "%"
        
        'Total Stock Volume
        Cells(j, 12).Value = T_Volume
        
        'Variable Resets
        T_Volume = 0
        Open_Value = Cells(i + 1, 3).Value
        
        'Ticker Row Iterator
        j = j + 1
    End If
    
Next i

End Sub


