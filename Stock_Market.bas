Attribute VB_Name = "Module1"
Sub wallStreet()

For Each ws In Worksheets
ws.Activate

'setup table for output
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

'find number of rows of data
Dim nRows As Long
nRows = Cells(Rows.Count, 1).End(xlUp).Row

'variables for data
Dim ticker As String
Dim openP As Double
Dim totalVolume As Long
Dim symbolCount As Long

'get first ticker symbol
If (Range("A2") <> "") Then
    ticker = Range("A2")
    symbolCount = 2
    openP = Range("C2")
    'totalVolume = 0
    Range("L2") = 0
End If

'loop through the rows in the sheet
For i = 2 To nRows + 1
   If (ticker = Cells(i, 1)) Then 'we are still on the same ticker symbol
        'totalVolume = totalVolume + Cells(i, 7)
        'volume = Cells(i, 7)
        Cells(symbolCount, 12) = Cells(symbolCount, 12) + Cells(i, 7)
    Else 'we are looking at a different ticker symbol
        'output previous data to the output table
        Cells(symbolCount, 9) = ticker
        Cells(symbolCount, 10) = Cells(i - 1, 6) - openP 'calculate and output yearly change
            'color cells green if yearly change is positive, red if negative
            If Cells(symbolCount, 10) > 0 Then
                Cells(symbolCount, 10).Interior.ColorIndex = 4
            ElseIf Cells(symbolCount, 10) < 0 Then
                Cells(symbolCount, 10).Interior.ColorIndex = 3
            End If
        If openP <> 0 Then
            Cells(symbolCount, 11) = Format(Cells(symbolCount, 10) / openP, "Percent") 'calculate and output percent change
            Else
                Cells(symbolCount, 11) = 0
        End If
        'Cells(symbolCount, 12) = totalVolume 'output stock volume
        'reset for new ticker symbol
        symbolCount = symbolCount + 1
        ticker = Cells(i, 1)
        openP = Cells(i, 3)
        'totalVolume = Cells(i, 7)
    End If
Next i

Next ws

End Sub


