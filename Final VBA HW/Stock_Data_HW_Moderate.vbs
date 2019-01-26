VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_HW_Moderate()
 
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
    
    'Set variables for brginning price, ending price, year change and percent change.
    'Also set initial value for beginning price on first stock
    Dim Yr_Begin_Price As Double
    Dim Yr_End_Price As Double
    Dim Yr_Change As Double
    Dim Percent_Change As Double
    Yr_Begin_Price = Cells(2, 3).Value
        
    ' Set variable to call summary row headings
    Dim Summary_Heading1 As String
    Dim Summary_Heading2 As String
    Dim Summary_Heading3 As String
    Dim Summary_Heading4 As String
    
    ' print summary table headings
    Summary_Heading1 = "Ticker"
    Summary_Heading2 = "Year Change"
    Summary_Heading3 = "Percent Change"
    Summary_Heading4 = "Total Stock Volume"
    
    ' Keep track of the location for each symbol in the summary table
    Dim Summary_Table_Row As LongPtr
    Summary_Table_Row = 2
      
    
      
    ' Loop through all symbols
    For I = 2 To Last_Row

        ' Check if we are still within the same symbol, if it is not...
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

            ' Set the Ticker Symbol
            Ticker_Symbol = Cells(I, 1).Value
            
            ' Add to the Stock Volume
             Stock_Volume = Stock_Volume + Cells(I, 7).Value
             
            'Select Year end price
            Yr_End_Price = Cells(I, 6).Value
            
            ' Calculate value for Year change
            Yr_Change = Yr_End_Price - Yr_Begin_Price
                
                'Check to see if percent change would require dividing by zero and just passing 0 to avoid error
                'Set formula for calculating percentage change
                If Yr_Begin_Price = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yr_Change / Yr_Begin_Price * 100
                End If

            ' Print the Summary Table
            Cells(Summary_Table_Row, 9).Value = Ticker_Symbol
            Cells(Summary_Table_Row, 12).Value = Stock_Volume
            Cells(Summary_Table_Row, 10).Value = Yr_Change
            Cells(Summary_Table_Row, 11).Value = Percent_Change & "%"
                             
            ' Reset the Stock Volume
            Stock_Volume = 0
            'Move to selecting the beginning price for the next ticker symbol
            Yr_Begin_Price = Cells(I + 1, 3).Value
                
                    'Set cell color for negative and positive values
                    If Cells(Summary_Table_Row, 10) > 0 Then
                        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    Else
                        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    End If
                    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
                     
        ' If the cell immediately following a row is the same symbol
        Else

        ' Add to the Volume Total
        Stock_Volume = Stock_Volume + Cells(I, 7).Value
                     
        End If

    Next I
   
  'Print the row headings for the summary table
  Cells(1, 9).Value = Summary_Heading1
  Cells(1, 10).Value = Summary_Heading2
  Cells(1, 11).Value = Summary_Heading3
  Cells(1, 12).Value = Summary_Heading4
  
 'Move to next worksheet
 Next ws
 
End Sub

