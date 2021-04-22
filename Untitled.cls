VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stonks():

Dim Ticker As LongLong 'Ticker is the ticker counter
Dim RowCount As LongLong  'rowcount will generally hold row # for reference
Dim LastRow As LongLong
Dim WS As Worksheet
Dim TickerName As String 'counter for ticker symbol only
Dim TickerVol As LongLong
Dim SumTableRow As LongLong
Dim ClosePrice As LongLong
Dim OpenPrice As LongLong
Dim FirstOpen As LongLong



'Loop through all sheets
For Each WS In Worksheets
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row


'Set up summary table
SumTableRow = 2
WS.Cells(1, 9).Value = "Ticker Name"
WS.Cells(1, 11).Value = "Yearly Price Change"
WS.Cells(1, 12).Value = "% Price Change"
WS.Cells(1, 10).Value = "Total Stock Volume"
WS.Range("A1:M1").Font.Bold = True

TickerVol = 0

    'loop through all ticker symbols
For Ticker = 2 To LastRow
    For RowCount = 2 To LastRow
    'check if next row is equal to current ticker row
        If WS.Cells(Ticker + 1, 1).Value <> WS.Cells(Ticker, 1).Value Then
     
        'grab the current ticker symbol value
        TickerName = WS.Cells(Ticker, 1).Value
    
       'print symbol in summary table row
        WS.Cells(SumTableRow, 9).Value = TickerName
    
        'increment summary table row by 1
        SumTableRow = SumTableRow + 1
    
      'sum of tickervol in summary table
        TickerVol = TickerVol + Cells(Ticker, 7).Value
     'print totalstockvolume to column 10
        WS.Cells(SumTableRow, 10).Value = TickerVol
    
    'increment TSG by 1
     SumTableRow = SumTableRow + 1
    
    'Reset TSV
     TickerVol = 0
     
   
        
    'Find close price at last row of ticker name
        ClosePrice = WS.Cells(Ticker - 1, 6).Value
        'Cells(4, 13).Value = ClosePrice
    

    
  
        
    End If
        Next RowCount
    Next Ticker
    'MsgBox (Cells(Ticker, 1).Value)
Next WS
MsgBox ("Done")


End Sub

Sub SummaryTable():


'Loop through all sheets


  'Find open price in first row of each ticker
     For RowCount = 2 To LastRow
     
        If Cells(RowCount, 3).Value <> 0 Then
        OpenPrice = Cells(RowCount, 3).Value
        Cells(15, 3).Value = OpenPrice
        Exit For
        
        End If
        
    Next RowCount
    MsgBox ("Done")
End Sub


