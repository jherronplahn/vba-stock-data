      Sub AlphaTest()

'Loop through the worksheets

        Dim WS As Worksheet
        For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        
            'Determine the last row in the worksheet
            lastRow = WS.Cells(Rows.Count, 1).End(xlUp).row

            'Add column headings
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Total Stock Volume"
            
            'Create variables
            Dim tickerNamer As String
            Dim totalVolume As Double
            totalVolume = 0
            Dim i As Long
            
            'Set staring row
            Dim row As Long
            row = 2
            Dim column As Long
            column = 1
            
            'Loop through rows
            For i = row To lastRow
                
                'Check ticker symbol
                If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                
                'Set ticker name
                tickerName = Cells(i, column).Value
                Cells(row, column + 8).Value = tickerName
                
                'Set ticker total volume per ticker
                totalVolume = totalVolume + Cells(i, column + 6).Value
                Cells(row, column + 9).Value = totalVolume
                
                'Set next row for ticker and total volume
                row = row + 1
                
                'Reset total volume per ticker to zero
                totalVolume = 0
                
                'if cells are the same ticker name
                Else
                totalVolume = volume + Cells(i, column + 6).Value
                
                End If
                
                
                Next i
        
         Next WS

      End Sub

