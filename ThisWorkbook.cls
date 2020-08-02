VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Main():

' Dim all variables used
Dim ticker As String
Dim tickers As Integer
Dim lastRowState As Long
Dim opening_price, closing_price, yearly_change, percent_change, totalvol As Double
Dim max_increase, max_decrease, max_vol As Double
Dim maxTickervol, maxTickerIncrease, maxTickerDecrease As String



' loop over each worksheet in the workbook
For Each ws In Worksheets

    
    ws.Activate

    ' Find the last row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Initialize variables
    tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    totalvol = 0
    
    
    For i = 2 To lastRowState

        ' Get the value of the ticker symbol we are currently calculating for.
        ticker = Cells(i, 1).Value
        
        ' Get the start of the year opening price for the ticker.
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Add up the total stock volume values for a ticker.
        totalvol = totalvol + Cells(i, 7).Value
        
        ' Run this if we get to a different ticker in the list.
        If Cells(i + 1, 1).Value <> ticker Then
            ' Increment the number of tickers when we get to a different ticker in the list.
            tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            
            closing_price = Cells(i, 6)
            yearly_change = closing_price - opening_price
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' Format Cells red and green to represent change
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
          
            End If
            
            
            ' Calculate percent change value for ticker.
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Format the percent_change value as a percent.
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            opening_price = 0
            
            ' Add to total volume
            Cells(number_tickers + 1, 12).Value = totalvol
            
            ' reset total volume
            totalvol = 0
        End If
        
    Next i
    'add headers
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Get the last row again
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables
    max_increase = Cells(2, 11).Value
    maxTickerIncreas = Cells(2, 9).Value
    max_decrease = Cells(2, 11).Value
    maxTickerDecrease = Cells(2, 9).Value
    max_volume = Cells(2, 12).Value
    maxTickervol = Cells(2, 9).Value
    
    
    
    For i = 2 To lastRowState
    
        ' Find ticker with highest increase, decrease, and volume
        If Cells(i, 11).Value > max_increase Then
            max_increase = Cells(i, 11).Value
            maxTickerIncrease = Cells(i, 9).Value
        End If

        If Cells(i, 11).Value < max_decrease Then
            max_decrease = Cells(i, 11).Value
            maxTickerDecrease = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > max_volume Then
           max_volume = Cells(i, 12).Value
           maxTickervol = Cells(i, 9).Value
        End If
        
    Next i
    
    Range("P2").Value = Format(maxTickerIncrease, "Percent")
    Range("Q2").Value = Format(max_increase, "Percent")
    Range("P3").Value = Format(maxTickerDecrease, "Percent")
    Range("Q3").Value = Format(max_decrease, "Percent")
    Range("P4").Value = maxTickervol
    Range("Q4").Value = max_volume
    
Next ws


End Sub
