Sub tickertotaler_easy()



'define variables
Dim ws As Worksheet
Dim ticker As String
Dim vol As Long
Dim Summary_Table_Row As Integer

On Error Resume Next

'loop through each worksheet
For Each ws In ThisWorkbook.Worksheets
    
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"

    Summary_Table_Row = 2

    'loop
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'find values
            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value


            'insert values
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1

            
      vol = 0
            Else
            vol = vol + ws.Cells(i, 7).Value
        
        End If
      
    Next i
    
Next ws


End Sub

