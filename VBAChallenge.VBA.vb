Sub stock()
Dim ticker As Integer
Dim percents As Double
Dim volume As Double
Dim tickerIndex As Long
Dim firsttime As Boolean
Dim openamt As Double
Dim closeamt As Double
Dim diffamt As Double
firsttime = True
tickerIndex = 1
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

For i = 2 To 797711
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    closeamt = Cells(i, 6)
    diffamt = closeamt - openamt
    tickerIndex = tickerIndex + 1
    Cells(tickerIndex, 9).Value = Cells(i, 1).Value
    Cells(tickerIndex, 10).Value = diffamt
    firsttime = True
    percents = (closeamt - openamt) / (openamt + 0.0000001)
    Cells(tickerIndex, 11).Value = percents
    volume = 0
Else
volume = volume + Cells(i, 7)
Range("L" & tickerIndex).Value = volume + Cells(i, 7)



End If

    If firsttime Then
        openamt = Cells(i, 3)
        firsttime = False
 End If
    
    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    ElseIf Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    Else
    Cells(i, 10).Interior.ColorIndex = 0
    End If
    
   


Next i
End Sub