Attribute VB_Name = "Module1"
Sub Stock()

    ' Loop through all sheets
    For Each ws In Worksheets

      'Populate header cells for new tables
      'Putting "ws." before each "Range" ensures the cells are filled on each sheet
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      ' Set initial variables for holding the Ticker symbol and total stock volume
      Dim Ticker As String
      Dim Total_Vol As Double
      
      ' Set initial variables for holding the Greatest Increases, Decreases, and Totals
      Dim Greatest_Increase As Double
      Dim Greatest_Decrease As Double
      Dim Greatest_Total As Double
      Dim Greatest_Increase_Ticker As Variant
      Dim Greatest_Decrease_Ticker As Variant
      Dim Greatest_Total_Ticker As Variant
      
      ' Set initial variables for holding the Ticker/Percent Change/Total Ranges
      ' for Index/Match Function and later formatting
      Dim Ticker_Range As Range
      Set Ticker_Range = ws.Range("I2:I2836")
      Dim Percent_Range As Range
      Set Percent_Range = ws.Range("K2:K2836")
      Dim Total_Range As Range
      Set Total_Range = ws.Range("L2:L2836")

      ' Set an initial variable for holding the last row in the data set and then find it
      Dim Last_Row As Long
      Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      ' Initialize Total_Vol to 0 before calculating
      Total_Vol = 0
    
      ' Keep track of the location for each Ticker symbol in *the summary table*
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
          
      ' Set initial variables for the open and close prices of each ticker
      Dim open_price As Double
      Dim close_price As Double
     
     ' identify the first open price before iterating through loop
     open_price = ws.Cells(2, 3).Value
     
      ' Loop through all stock amounts
      For i = 2 To Last_Row
      
        ' Add to the Volume Total
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        
        ' Begin condition that identifies a change in ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
          ' Set the Ticker Symbol
          Ticker = ws.Cells(i, 1).Value
    
          ' Print the Ticker Symbol in the Summary Table
          ws.Range("I" & Summary_Table_Row).Value = Ticker
    
          ' Print the Volume Total to the Summary Table
          ws.Range("L" & Summary_Table_Row).Value = Total_Vol
          
          ' Define the close price according to cell values at each iteration of ticker change
          close_price = ws.Cells(i, 6).Value
    
          ' Set initial variables for the yearly change and percent change for each ticker
          Dim yearly_change As Double
          Dim percent_change As Double
         
         ' Calculate yearly_change and percent_change
          yearly_change = close_price - open_price
          
         ' To avoid an Overflow error, we must check if open_price is ever equal to 0
         ' so as to never divide by 0
          If open_price <> 0 Then

            percent_change = yearly_change / open_price
          
            Else: percent_change = 0
          
          End If
          
          ' Define the open and close prices according to cell values at each iteration of ticker change
          open_price = ws.Cells(i + 1, 3).Value
    
         ' Print the Yearly change in opening and closing price to the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = yearly_change
         
         ' Print the % Change in opening and closing price for the year to the Summary Table
          ws.Range("K" & Summary_Table_Row).Value = percent_change
          
          ' Format the colors of the yearly change column
          If yearly_change < 0 Then
        
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
            Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
          
          End If
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Ticker Total
          Total_Vol = 0
    
        End If
     
     ' Iterate to next Ticker Symbol
      Next i
      
    ' VBA macro for finding max and min values
    Greatest_Increase = Application.Max(Percent_Range)
    Greatest_Decrease = Application.Min(Percent_Range)
    Greatest_Total = Application.Max(Total_Range)
    
    ' VBA macro for Index/Match to associate Tickers with min/max/total values
    Greatest_Increase_Ticker = WorksheetFunction.Index(Ticker_Range, WorksheetFunction.Match(Greatest_Increase, Percent_Range, 0))
    Greatest_Decrease_Ticker = WorksheetFunction.Index(Ticker_Range, WorksheetFunction.Match(Greatest_Decrease, Percent_Range, 0))
    Greatest_Total_Ticker = WorksheetFunction.Index(Ticker_Range, WorksheetFunction.Match(Greatest_Total, Total_Range, 0))

    ' Populate Table with results from Max/Min calculations
    ws.Range("Q2").Value = Greatest_Increase
    ws.Range("Q3").Value = Greatest_Decrease
    ws.Range("Q4").Value = Greatest_Total

    ' Populate Table with results from Index/Match
    ws.Range("P2").Value = Greatest_Increase_Ticker
    ws.Range("P3").Value = Greatest_Decrease_Ticker
    ws.Range("P4").Value = Greatest_Total_Ticker
    
    'Format Cells
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Percent_Range.NumberFormat = "0.00%"

' Iterate to next worksheet in the workbook
Next ws

' Just have this here to confirm script is done running
MsgBox ("Subroutine Complete")

End Sub



