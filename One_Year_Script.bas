Attribute VB_Name = "Module1"
Sub YearStock()
Dim Open_Price As Double
Open_Price = Cells(2, 3)
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Ticker_Name As String
Dim Percent_Change As Double
Dim Volume As Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim increase_number As Double
Dim decrease_number As Double
Dim volume_number As Double




LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_Name = (Cells(i, 1).Value)
Close_Price = Cells(i, 6).Value
Yearly_Change = Close_Price - Open_Price
Percent_Change = (Yearly_Change / Open_Price)
Range("I" & Summary_Table_Row).Value = Ticker_Name
Range("J" & Summary_Table_Row).Value = Yearly_Change
If Yearly_Change > 0 Then
    Range("J" & Summary_Table_Row).Interior.Color = vbGreen
    Else
    Range("J" & Summary_Table_Row).Interior.Color = vbRed
    End If
Range("K" & Summary_Table_Row).Value = Percent_Change
If Percent_Change > 0 Then
Range("K" & Summary_Table_Row).Interior.Color = vbGreen
Else
Range("K" & Summary_Table_Row).Interior.Color = vbRed
End If
Summary_Table_Row = Summary_Table_Row + 1
Open_Price = Cells(i + 1, 3)
End If
Volume = Volume + Cells(i, 7).Value
Range("L" & Summary_Table_Row).Value = Volume
Next i


    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

End Sub
