VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub stock_analysis()

    'Create variables
    Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
    Dim last_Row As Long
    Dim ticker_Symbol As String
    Dim quarterly_Change As Double
    Dim percent_Change As Double
    Dim total_Volume As Double
    Dim open_Price As Double
    Dim close_Price As Double
    Dim row_Start As Long
    Dim greatest_Increase As Double
    Dim greatest_Decrease As Double
    Dim greatest_Volume As Double
    Dim tickerGreatest_Increase As String
    Dim tickerGreatest_Decrease As String
    Dim tickerGreatest_Volume As String
    Dim row As Double
    
    
    ' Initialize variables
    last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    greatest_Increase = 0
    greatest_Decrease = 0
    greatest_Volume = 0
    total_Volume = 0
    row_Start = 2
    
    ' Create Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greates Total Volume"
    
    'Begin the loop
    For row = 2 To last_Row
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                ticker_Symbol = ws.Cells(row, 1).Value
                total_Volume = total_Volume + ws.Cells(row, 7).Value
                close_Price = ws.Cells(row, 6).Value
               quarterly_Change = close_Price - open_Price
               If open_Price <> 0 Then
                    percent_Change = (quarterly_Change / open_Price) * 100
        Else
            percent_Change = 0
        End If
        
        ' Print output
        ws.Range("I" & row_Start).Value = ticker_Symbol
        ws.Range("J" & row_Start).Value = quarterly_Change
        ws.Range("K" & row_Start).Value = percent_Change
        ws.Range("L" & row_Start).Value = total_Volume
        
        ' Apply Conditional format
        If quarterly_Change >= 0 Then
            ws.Range("J" & row_Start).Interior.Color = vbGreen
            Else
                ws.Range("J" & row_Start).Interior.Color = vbRed
        End If
        
        ' Check and update greatest increase, decrease, and volume
        If percent_Change > greatest_Increase Then
            greatest_Increase = percent_Change
            tickerGreatest_Increase = ticker_Symbol
                ElseIf percent_Change < greatest_Decrease Then
                    greatest_Decrease = percent_Change
                    tickerGreatest_Decrease = ticker_Symbol
        End If
        
        If total_Volume > greatest_Volume Then
            greatest_Volume = total_Volume
            tickerGreatest_Volume = ticker_Symbol
            
        End If
        
            total_Volume = 0
            row_Start = row_Start + 1
            open_Price = ws.Cells(row + 1, 3).Value
        Else
            total_Volume = total_Volume + ws.Cells(row, 7).Value
            
    End If
    
Next row
        ' Output the greatest increase, decrease, and volume
        ws.Range("p2").Value = tickerGreatest_Increase
        ws.Range("q2").Value = greatest_Increase
        ws.Range("p3").Value = tickerGreatest_Decrease
        ws.Range("q3").Value = greatest_Decrease
        ws.Range("p4").Value = tickerGreatest_Volume
        ws.Range("q4").Value = greatest_Volume

Next ws

End Sub



