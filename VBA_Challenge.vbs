Attribute VB_Name = "Module1"
Sub YearlyStocks():

    Dim ws As Worksheet
    Dim year As String
    Dim ticker As String
    Dim yearlychange As Double
    Dim percent As Double
    Dim volume As Variant
    
    'Loop through each worksheet.
    For Each ws In ThisWorkbook.Worksheets
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
        For i = 2 To lastrow
            
            'Loop through each row until next ticker.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
            'Determine number of stock days to use.
                year = Left(ws.Range("B2"), 4)
                If year = "2020" Then
                    firstcell = i - 252
                Else
                    firstcell = i - 251
                End If
    
            'Pull each ticker.
               ticker = ws.Cells(i, 1).Value
               ws.Range("I" & Summary_Table_Row).Value = ticker
               
            'Calculate yearly change (253 open stock days in 2020) and highlight in red if negative or green if positive.
               yearlychange = (ws.Cells(i, 6).Value - ws.Cells(firstcell, 3).Value)
               ws.Range("J" & Summary_Table_Row).Value = yearlychange
               If yearlychange < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
        
            'Calculate percent change (yearly change/opening price) and format to a percentage.
               percent = (ws.Cells(i, 6).Value - ws.Cells(firstcell, 3).Value) / (ws.Cells(firstcell, 3).Value)
               ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percent)
            
            'Calculate total stock volume.
               volume = Application.Sum(ws.Range(ws.Cells(firstcell, 7), ws.Cells(i, 7)))
               ws.Range("L" & Summary_Table_Row).Value = volume
            
            'Add one to the summary table row
               Summary_Table_Row = Summary_Table_Row + 1
            Else
                
            End If
            
        Next i
        
        'Create columns for Ticker, Yearly Change, Percent Change, Total Stock Volume
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
    
    Next ws

End Sub
