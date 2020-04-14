Attribute VB_Name = "Módulo1"

Sub Ticker_Analysis()

'Loop through all sheets
    Dim ws As Worksheet
    
    For Each ws In Worksheets
            ws.Activate
            
 'Variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    Dim volume As Double
    Dim R As Long
    Dim rowdelimiter As Integer
    Dim SumTabRow As Integer
    Dim MaxIncreaseTicker As String
    Dim MaxIncreaseValue As Double
    Dim MaxDecreaseTicker As String
    Dim MaxDecreaseValue As Double
    Dim MaxVolumeTicker As String
    Dim MaxVolumeValue As Double
    
                 volume = 0
                rowdelimiter = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Store information
        SumTabRow = 2
            
'Headers
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Open Price"
            Range("K1").Value = "Close Price"
            Range("L1").Value = "Volume"
            Range("M1").Value = "Yearly Change"
            Range("N1").Value = "Percent Change"

            
'Iterate
            For i = 2 To rowdelimiter
                        
                        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                                ticker = Cells(i, 1).Value
                                    Range("I" & SumTabRow).Value = ticker
                                
                                openPrice = Cells(i, 3).Value
                                    Range("J" & SumTabRow).Value = openPrice
                                 
                                 closePrice = Cells(i, 6).Value
                                    Range("K" & SumTabRow).Value = closePrice
                                
                                volume = Cells(i, 7).Value
                                    Range("L" & SumTabRow).Value = volume
                                
                                yearlyChange = closePrice - openPrice
                                    Range("M" & SumTabRow).Value = yearlyChange
                                
                                percentChange = (yearlyChange / openPrice)
                                    Range("N" & SumTabRow).Value = percentChange
                                    Range("N" & SumTabRow).NumberFormat = "0.00%"
                                    
                                    
                                    
                                    Range("L:N").ColumnWidth = 15
                                    
                                SumTabRow = SumTabRow + 1
                                   
                        End If

                                    
 
            Next i


            
    Next ws
                        
End Sub

    Sub reset()

Dim ws As Worksheet
    
    For Each ws In Worksheets
    ws.Activate
    
    Range(["I:Z"]).Delete
    
    
    Next ws

End Sub

 

Sub winnersAndLosers()

Dim Max As Double
Dim ws As Worksheet
    For Each ws In Worksheets
ws.Activate

'LayOut Winners & Losers
    ws.Cells(2, 18).Value = "Greatest % Increase"
    ws.Cells(3, 18).Value = "Greatest & Decrease"
    ws.Cells(4, 18).Value = "Greatest Total Volume"
    ws.Cells(1, 19).Value = "Ticker"
    ws.Cells(1, 20).Value = "Value"
 Range("R:T").ColumnWidth = 20

'Print Results

' For Value
    ws.Cells(2, 20).Value = Application.WorksheetFunction.Max(Columns(13))
    ws.Cells(3, 20).Value = Application.WorksheetFunction.Min(Columns(13))
    ws.Cells(4, 20).Value = Application.WorksheetFunction.Min(Columns(12))
    
    'For Tickers
    ws.Cells(2, 19).Value = Application.WorksheetFunction.Max(Columns(13))
    ws.Cells(3, 19).Value = Application.WorksheetFunction.Min(Columns(13))
    ws.Cells(4, 19).Value = Application.WorksheetFunction.Min(Columns(12))

Next ws
End Sub

