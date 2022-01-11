Attribute VB_Name = "Module1"
Sub stocks()

    For Each ws In Worksheets
        ' get number of rows
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim totalvol As LongLong
        totalvol = 0
        Dim row As Long
        row = 2
        Dim openprice As Double
        openprice = ws.Cells(2, 3).Value
        Dim closeprice As Double
        
        Dim i As Long
        ' loop through columns
        For i = 2 To lastrow
        
        totalvol = totalvol + ws.Cells(i, 7).Value
            
            ' check if new stock next
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                closeprice = ws.Cells(i, 6).Value
                
            
                ws.Cells(row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(row, 10).Value = openprice - closeprice
                ws.Cells(row, 11).NumberFormat = "0.00%"
                If openprice <> 0 Then
                    ws.Cells(row, 11).Value = (closeprice - openprice) / openprice
                End If
                ws.Cells(row, 12).Value = totalvol
                totalvol = 0
                
                openprice = ws.Cells(i + 1, 3).Value
                
                row = row + 1
                
            End If
        
        Next i
        
        Dim rownum As Long
        rownum = ws.Cells(Rows.Count, 10).End(xlUp).row
        For j = 2 To rownum
            If ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
                
            ElseIf ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            End If
        Next j
            
            
    
    Next ws

End Sub


