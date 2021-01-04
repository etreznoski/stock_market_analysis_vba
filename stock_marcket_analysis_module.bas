Attribute VB_Name = "Module1"
'identify/capture open and close values
'yearly change = close-open
'percent change = yearly change / open
'display percent change as a percent



        
Sub stock_data()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Name columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


'Set variable to hold ticker name
Dim ticker As String

'Set other variables
Dim open_price As Double

open_price = Range("C2").Value

Dim close_price As Double
'Dim yearly_change As Double

Dim ticker_volume As LongLong
ticker_volume = 0

'Keep track of location of each ticker name in summary table
Dim summary_row As Long
summary_row = 2

Dim rowcount As Double
rowcount = Cells(Rows.Count, 1).End(xlUp).Row


'Loop through all ticker names
For i = 2 To rowcount

    'check if we are still in the same ticker name, if it is not..
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'set ticker name
        ticker = Cells(i, 1).Value
        
        'add to ticker volume total
        ticker_volume = ticker_volume + Cells(i, 7).Value
        
        close_price = Cells(i, 6).Value
        
        Range("J" & summary_row).Value = close_price - open_price
        
        Range("k" & summary_row).Value = (close_price - open_price) / open_price
        
        open_price = Cells(i + 1, 3).Value
        
        'print ticker name in summary table
        Range("I" & summary_row).Value = ticker
        
        'print ticker volume in summary table
        Range("L" & summary_row).Value = ticker_volume
        
        ''print opening price in summary table
        'Range("J" & summary_row).Value = open_price
        

        'add one to the summary table row
        summary_row = summary_row + 1
        
        'reset volume
        ticker_volume = 0
        
        'if still in same ticker
        Else
        
            'add to volume
            ticker_volume = ticker_volume + Cells(i, 7).Value
            
            
    End If


Next i

end_summary = Cells(Rows.Count, 10).End(xlUp).Row
'change formatting to percent
Range("K2" & end_summary).NumberFormat = "0.00%"


'conditional formating to color green or red based on positive or nagative yearly change
'find end of summary table row


For j = 2 To end_summary
    If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    Else
        Cells(j, 10).Interior.ColorIndex = 3
    
    'this will also color cells with zero value red..
    End If
Next j



'loop through percent change and store greatest value,
'replacing each time higher value is found
    'store corresponding ticker
    'print to worksheet
        'repeat with lowest value
        
'set titles
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    Dim greatest_change As Double
    Dim ticker_one As String
 
    greatest_change = Cells(2, 11).Value
    
    For k = 2 To end_summary
         If Cells(k, 11).Value > greatest_change Then
         
             greatest_change = Cells(k, 11).Value
            ticker_one = Cells(k, 9).Value
        
          End If
       Range("N2").Value = "Greatest % Increase"
       Range("O2").Value = ticker_one
       Range("P2").Value = greatest_change
        
    Next k
    'change formatting of cell P2 to percent**************
    Range("P2").NumberFormat = "0.00%"
    
        Dim low_change As Double
        Dim ticker_two As String
        
    low_change = Cells(2, 11).Value
    
    For m = 2 To end_summary
    
     If Cells(m, 11).Value < low_change Then
            low_change = Cells(m, 11).Value
            ticker_two = Cells(m, 9).Value
            
        End If
   
        Range("N3").Value = "Greatest % Decrease"
        Range("O3").Value = ticker_two
        Range("P3").Value = low_change
     Next m
     'change formatting of cell P3 to percent**************
    Range("P3").NumberFormat = "0.00%"
    
'loop through the total volume, store greatest value, replacing when higher value found
    'store corresponding ticker
    'print to worksheet
        Dim greatest_volume As LongLong
        Dim ticker_three As String
        
    greatest_volume = Cells(2, 12).Value
    
    For n = 2 To end_summary
    
     If Cells(n, 12).Value > greatest_volume Then
            greatest_volume = Cells(n, 12).Value
            ticker_three = Cells(n, 9).Value
            
        End If
   
        Range("N4").Value = "Greatest Total Volume"
        Range("O4").Value = ticker_three
        Range("P4").Value = greatest_volume
     Next n


End Sub


