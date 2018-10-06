VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_year_stock_data()
For Each ws In Worksheets
ws.Activate
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total_Stock_Volume"
    Dim Ticker_Name As String
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_Name = Cells(i, 1).Value
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Range("J" & Summary_Table_Row).Value = Total_Stock_Volume
            Summary_Table_Row = Summary_Table_Row + 1
            Total_Stock_Volume = 0
        Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        End If
    
    Next i

Next ws

End Sub
