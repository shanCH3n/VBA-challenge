Attribute VB_Name = "Module1"

Sub StockTally()
    For Each ws In Worksheets
        Dim i As Long
    
        Dim ticker_name As String
        Dim YrChg As Double
        Dim PerChg As Double
  

        Dim TolStkVol As LongLong
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

        Dim Start As Long 'To track first occurrence of stock
        Start = 2
    
        'Formatting of summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("J").ColumnWidth = 14
        ws.Columns("K").ColumnWidth = 14
        ws.Columns("L").ColumnWidth = 18
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("K").NumberFormat = "0.00%"

    
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                YrChg = ws.Cells(i, 6) - ws.Cells(Start, 3)
                PerChg = (YrChg / ws.Cells(Start, 3).Value)
                TolStkVol = TolStkVol + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                ws.Range("J" & Summary_Table_Row).Value = YrChg
                ws.Range("K" & Summary_Table_Row).Value = PerChg
                ws.Range("L" & Summary_Table_Row).Value = TolStkVol
                Start = i + 1
                Summary_Table_Row = Summary_Table_Row + 1
                TolStkVol = 0
            
            Else
                TolStkVol = TolStkVol + ws.Cells(i, 7).Value
              
            End If
        Next i
    
    'Add conditional formatting to YrChg (Column J)
    
        Dim SumLastRow As Long
        SumLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To SumLastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Cells(i, 10).Interior.Color = RGB(255, 255, 255)
            End If
        Next i

    Next ws
    
End Sub
