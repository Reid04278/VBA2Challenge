Attribute VB_Name = "Module2"
Sub WorkSheetStockAnalyzer()
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
Dim wbk As Workbook

currpath = ActiveWorkbook.Path
Set wbk = Workbooks.Open(currpath & "\multiple_year_stock_data.xlsx")

For Each ws In wbk.Worksheets





LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker_Name = (ws.Cells(i, 1).Value)
Close_Price = ws.Cells(i, 6).Value
Yearly_Change = Close_Price - Open_Price
If Open_Price <> 0 Then
Percent_Change = (Close_Price - Open_Price) / Open_Price
Else
Percent_Change = 0
End If

ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
If Yearly_Change > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
    Else
    ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
    End If
ws.Range("K" & Summary_Table_Row).Value = Percent_Change
If Percent_Change > 0 Then
ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
Else
ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed
End If
Summary_Table_Row = Summary_Table_Row + 1
Open_Price = ws.Cells(i + 1, 3)
End If
Volume = Volume + ws.Cells(i, 7).Value
ws.Range("L" & Summary_Table_Row).Value = Volume
Next i


    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)
    Next ws
    
    
    
    

End Sub
