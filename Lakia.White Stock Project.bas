Attribute VB_Name = "Module1"
Sub StockAnalysis():
For Each ws In Worksheets
Dim TickerName As String

Dim OpenPrice As Double
    OpenPrice = ws.Cells(2, "C")

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
Dim ClosePrice As Double

Dim Vol As Double
    Vol = 0
        ws.Cells(1, "L") = "Ticker"
        
        ws.Cells(1, "M") = "Annual Change"
        
        ws.Cells(1, "N") = "Percentage Change"
        
        ws.Cells(1, "O") = "Total Volume"
        
        ws.Cells(1, "S") = "Ticker"
        
        ws.Cells(1, "T") = "Value"
        
        ws.Cells(2, "R") = "Greatest % increase"
        
        ws.Cells(3, "R") = "Greatest % decrease"
        
        ws.Cells(4, "R") = "Greatest total volume"
        

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

        Vol = Vol + ws.Cells(i, "G")
        
If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then

    ws.Cells(Summary_Table_Row, "L") = ws.Cells(i, 1).Value
    
    ClosePrice = ws.Cells(i, "F")
        ws.Cells(Summary_Table_Row, "M") = ClosePrice - OpenPrice
    
    ws.Cells(Summary_Table_Row, "O") = Vol
    
If OpenPrice <> 0 Then
    
        ws.Cells(Summary_Table_Row, "N") = FormatPercent((ClosePrice - OpenPrice) / OpenPrice, 2)
    
    Else
    
        ws.Cells(Summary_Table_Row, "N") = 0
    
    End If
    
    
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    OpenPrice = ws.Cells(i + 1, "C")
    Vol = 0
    
End If

Next i
    Next ws

End Sub
