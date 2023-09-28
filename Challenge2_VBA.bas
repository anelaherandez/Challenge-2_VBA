Attribute VB_Name = "Challenge2_VBA"
Sub stockvolume()

For Each ws In ThisWorkbook.Worksheets

'Declare all variables
    Dim TickerSymbol As String
    Dim TotalStockVolume As Long
    Dim ws As Worksheet
    
    
'create columns
    ws.Range("J1") = "Ticker Symbol"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
    ws.Range("P2") = "Greatest % Increase"
    ws.Range("P3") = "Greatest % Decrease"
    ws.Range("P4") = "Greatest Total Volume"
    ws.Range("Q1") = "Ticker"
    ws.Range("R1") = "Value"
    
'Designate Last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Designate the variable name and where it will start that will track the volume amount
    Total_Volume = 0
    
'Designate the name for the summary chart of Ticker names
    SummaryTS_Chart = 2
    
'Begin the Loop to evaluate the volume total for each ticker name
    For input_row = 2 To LastRow
    
        If ws.Cells(input_row + 1, 1).Value <> ws.Cells(input_row, 1) Then
        'set ticker name
            Ticker_Name = ws.Cells(input_row, 1).Value
        'Add to the total volume
            Total_Volume = Total_Volume + ws.Cells(input_row, 7).Value
        'Print Ticker name in summary column
            ws.Range("J" & SummaryTS_Chart).Value = Ticker_Name
            ws.Range("M" & SummaryTS_Chart).Value = Total_Volume
        'Add 1 to SummaryTS_Chart
            SummaryTS_Chart = SummaryTS_Chart + 1
        'Restart Volume tracker
            Total_Volume = 0
        
        Else
            Total_Volume = Total_Volume + ws.Cells(input_row, 7).Value
    End If
    Next input_row
    
    ws.Columns("A:R").AutoFit
    
    'Exit For
            
Next ws


End Sub
Sub MaxVolume()

'Declare all variable
Dim LastRow As Long


For Each ws In ThisWorkbook.Worksheets

ws.Activate

'Designate the last row to analyze the greatest stock volume

    LastRow = ws.Range("M" & Rows.Count).End(xlUp).Row
    
'Use the max function to find the greatest total volume for each worksheet

     ws.Range("R4") = WorksheetFunction.Max(Range("M2:M" & LastRow))
     
'Use the xlookup function to find the corresponding Ticker symbol

    ws.Range("Q4") = WorksheetFunction.XLookup([R4], [M:M], [J:J])
 
 Next ws
 
 
End Sub

Sub YearlyChange()

'Declare all variables
    Dim TickerSymbol As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    
For Each ws In ThisWorkbook.Worksheets
  'Designate LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
 'Designate the Summary Chart for Tickernames
        SummaryTS_Chart = 2
        
 'Designate the beginning of the open price
        Price_Row = 2
        
 'Loop Through the worksheet to calculate the yearly change and percentage change for each ticker symbol
 
For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        
        'Designate the ticker symbol
            Ticker_Symbol = Cells(i, 1).Value
         'Specify where to look for the opening price for each ticker
         
            Open_Price = ws.Range("C" & Price_Row).Value
        'Specify where to look for the closing price for each ticker
        
            Close_Price = ws.Range("F" & i).Value
        'Calculate the Yearly Change
        
            YearlyChange = Close_Price - Open_Price
        'Calculate the Percent Change
        
            PercentChange = YearlyChange / Open_Price
        'Populate the Columns with Yearly and Percent Change as well as format the percent change
        
            ws.Range("K" & SummaryTS_Chart).Value = YearlyChange
            ws.Range("L" & SummaryTS_Chart).Value = PercentChange
            ws.Range("L" & SummaryTS_Chart).NumberFormat = "0.00%"
            
        'Format the Yearly Change column to change color whether positive or negative
        
            
                 If ws.Range("K" & SummaryTS_Chart).Value > 0 Then
                    ws.Range("K" & SummaryTS_Chart).Interior.ColorIndex = 4
                 Else
                    ws.Range("K" & SummaryTS_Chart).Interior.ColorIndex = 3
                 End If
                 
    'Add 1 to Summary chart for next loop
        
        SummaryTS_Chart = SummaryTS_Chart + 1
        
    'Add 1 to price row for next loop
    
         Price_Row = i + 1
   End If
         
Next i

Next ws

End Sub

Sub MaxMinPercentChange()

'Declare all variable
Dim LastRow As Long

For Each ws In ThisWorkbook.Worksheets

ws.Activate

'Designate Last Row
    LastRow = ws.Range("L" & ws.Rows.Count).End(xlUp).Row
    
'Find the Greatest Increase Of Percent Change Using Max Function
    
    ws.Range("R2") = WorksheetFunction.Max(Range("L2:L" & LastRow))
    
'Format the cell to be in percentage

    ws.Range("R2").NumberFormat = "0.00%"
    
'Find the corresponding Ticker Symbol using Xlookup Function

    ws.Range("Q2") = WorksheetFunction.XLookup([R2], [L:L], [J:J])
    
'Find the Greatest Decrease Of Percent Change Using Max Function

    ws.Range("R3") = WorksheetFunction.Min(Range("L2:L" & LastRow))
    
'Format the cell to be in percentage

    ws.Range("R3").NumberFormat = "0.00%"
    
'Find the corresponding Ticker Symbol using Xlookup Function

    ws.Range("Q3") = WorksheetFunction.XLookup([R3], [L:L], [J:J])
    
    
    
Next ws

    
End Sub

