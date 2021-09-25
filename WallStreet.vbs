Sub Runallsheets()
    Dim wSh As Worksheet
    Application.ScreenUpdating = False
    For Each wSh In Worksheets
        wSh.Select
        Call WallStreet
    Next
    Application.ScreenUpdating = True
End Sub

Sub WallStreet()

    Dim volumne_tot As LongLong
    Dim ticker As String
    Dim ticker_ctr_sum As Double
    Dim ticker_open_close_ctr As Double
    Dim yearly_open As Double
    Dim yearly_end As Double
    
    volumne_tot = 0
    yearly_end = 0
    yearly_open = 0
    
    'keep row ctr for ticker summary row
    ticker_ctr_sum = 2
    
    'keep row ctr for open and close value
    ticker_open_close_ctr = 2
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        volumne_tot = volumne_tot + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(ticker_open_close_ctr, 3)
 
        ' write summary table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'set price for year end close
            yearly_end = Cells(i, 6)
            
            'print ticker to sum table
            Cells(ticker_ctr_sum, 9).Value = ticker
            
            'print year change to sum table
            Cells(ticker_ctr_sum, 10).Value = yearly_end - yearly_open

            If yearly_open = 0 Then
                Cells(ticker_ctr_sum, 11).Value = Null
            Else
                Cells(ticker_ctr_sum, 11).Value = (yearly_end - yearly_open) / yearly_open
            End If
            
            'print total volume
            Cells(ticker_ctr_sum, 12).Value = volumne_tot
            
            ' Color the cell green if > 0, red if < 0
            If Cells(ticker_ctr_sum, 10).Value > 0 Then
                Cells(ticker_ctr_sum, 10).Interior.ColorIndex = 4
            Else
                Cells(ticker_ctr_sum, 10).Interior.ColorIndex = 3
            End If
            
            Cells(ticker_ctr_sum, 11).NumberFormat = "0.00%"
            
            volumne_tot = 0
            
            ticker_ctr_sum = ticker_ctr_sum + 1
            ticker_open_close_ctr = i + 1
            
        End If
        
    Next i

    Columns("J").AutoFit
    Columns("K").AutoFit
    Columns("L").AutoFit

End Sub



