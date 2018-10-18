Sub AddVolume()
    
    
    For Each ws In Worksheets

        Dim Ticker As String
        Dim TotalStockVolume As Double

        TotalStockVolume = 0

        
        
        sumTableRow = 2

        Dim LastRow As Long, i As Long



        Dim WorksheetName As String
            
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                For i = 2 To LastRow

                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                
                        Ticker = ws.Cells(i, 1).Value

                
                        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                        ws.Range("I" & sumTableRow).Value = Ticker

                
                        ws.Range("J" & sumTableRow).Value = TotalStockVolume
                
                        sumTableRow = sumTableRow + 1

                
                        TotalStockVolume = 0

            
                    Else
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                    End If
                Next i
                
       
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Total Stock Volume"
    Next
End Sub
