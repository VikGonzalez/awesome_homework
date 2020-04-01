Attribute VB_Name = "Módulo1"
Option Explicit

Sub stock_analysis()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

ws.Activate

    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    Dim volume As Double
    volume = 0
    
   Rowdelimitier = Cells(Rows.Count, 1).End(xlUp).Row
   
   open_price = Cells(2, 3).Value
   
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
   
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Value"
   
   For i = 2 To Rowdelimitier
   
        If Cells(1 + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value
        
        close_price = Cells(i, 6).Value
        
        yearlyChange = close_price - open_price
        percentChange = (yearlyChange / open_price)
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        volume = volume´ + Cells(i, 7).Value
        
    Range("I" & Summary_Table_Row).Value = "Ticker"
    Range("J" & Summary_Table_Row).Value = "Yearly Change"
    Range("K" & Summary_Table_Row).Value = "Percent Change"
    Range("L" & Summary_Table_Row).Value = "Total Stock Value"
   

End Sub
