VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Volume_HW_Easy()
 
 'Set variable for defining the worksheet
 Dim ws As Worksheet
 
 'Start Loop to execute script over entire workbook
 For Each ws In ThisWorkbook.Worksheets
    
    'Script to activate worksheet
    ws.Activate
    
    ' Set an initial variable for holding the ticker symbol
    Dim Ticker_Symbol As String
  
    'Declare last row as long and set formula for working to last row
    Dim Last_Row As Long
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Set an initial variable for holding the total volume
    Dim Stock_Volume As Double
    Stock_Volume = 0

    ' Set summary row headings
    Dim Summary_Heading1 As String
    Dim Summary_Heading2 As String
    
    Summary_Heading1 = "Ticker"
    Summary_Heading2 = "Total Stock Volume"
    
    ' Keep track of the location for each symbol in the summary table
    Dim Summary_Table_Row As LongPtr
    Summary_Table_Row = 2
      
    ' print summary table headings
      
    ' Loop through all symbols
    For I = 2 To Last_Row

        ' Check if we are still within the same symbol, if it is not...
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

            ' Set the Ticker Symbol
            Ticker_Symbol = Cells(I, 1).Value

            ' Add to the Stock Volume
             Stock_Volume = Stock_Volume + Cells(I, 7).Value

            ' Print the Ticker Symbol in the Summary Table
            Cells(Summary_Table_Row, 9).Value = Ticker_Symbol

            ' Print the Stock Volume to the Summary Table
            Cells(Summary_Table_Row, 10).Value = Stock_Volume

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Stock Volume
            Stock_Volume = 0

        ' If the cell immediately following a row is the same symbol
        Else

        ' Add to the Volume Total
        Stock_Volume = Stock_Volume + Cells(I, 7).Value

        End If

    Next I
    
  Cells(1, 9).Value = Summary_Heading1
  Cells(1, 10).Value = Summary_Heading2
    
 'MsgBox ws.Name
 
 Next ws
 
End Sub

