

'Sub to run the script on all worksheets
Sub OnAllSheets()

    'Set the variable for the worksheet
    Dim sh As Worksheet
    
    'Turn off screen updating so scirpt runs faster
    Application.ScreenUpdating = False
    
    'Activate, Select and call main subroutine to run on each sheet
    For Each sh In Worksheets
        sh.Activate
        sh.Select
        Call stockSummary
    Next
    
    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub


'Main subroutine for the assignment
Sub stockSummary()

    'Set the initial variables for creating the summary table
    Dim Current_Stock As String
    Dim Current_Open As Currency
    Dim Current_Close As Currency
    Dim Current_Percentage As Double
    Dim Total_Stock_Volume As Double

    
    'Setting variables for the Bonus Section
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Double

 
 '---SUMMARY TABLE -----
'Set column headers for Summary Table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"

'Keep track of the location in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
'Find the last row for the stock section
 LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
'Ensure the table is sorted first by name (Key1), then by date (Key2)
Range("A:G").Sort Key1:=Range("A1"), _
                     Order1:=xlAscending, _
                     Key2:=Range("B1"), _
                     Order2:=xlAscending, _
                     Header:=xlYes


'Loop through all of the rows
For I = 2 To LastRow

    'Is the stock in the current row equal to Current_Stock?  If not, it's a new stock that still needs to be added to the summary.
    'Current_Stock will be set to the stock in the current row so its details can be added to the summary
    If Cells(I, 1).Value <> Current_Stock Then
    
        'Setting current stock and current open to the value in this iteration.
        Current_Stock = Cells(I, 1).Value
        Current_Open = Cells(I, 3).Value
        
        'Populating the summary table with the current_stock.
        Range("I" & Summary_Table_Row).Value = Current_Stock
        
        'Resetting to total_stock_volume to 0 as this is the first entry for this stock's ticker.
        Total_Stock_Volume = 0
        
    'If current_stock is the same as the stock in this row then we need to continue iterating through to calculate the sum the volume and to find the close.
    Else
    
      'Add each new volume entry to Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
      
      'Populate the summary table with Total_Stock_Volume
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
         
       'Test to see if we've reached the end of the Current_Stock's data by adding one to the current row and comparing to Current_Stock.
       'If they are different, we've reached the end.
        If Cells(I + 1, 1).Value <> Current_Stock Then
        
            'Since we're at the end, the figure in the close column will be the final close for the year.  That's set to Current_Close.
            Current_Close = Cells(I, 6)
            
            'As there are some Open cells with a 0 value, there will be an Divide by Zero error when calculating percentages. This code moves to the next line rather than raising an error
            On Error Resume Next
            
            'Determining percentage by subtracting Current_Open from Current_Close and dividing by the Current_Open
            Current_Percentage = (Current_Close - Current_Open) / Current_Open
            
            'formating it to be percentage
            Current_Percentage_formatted = Format(Expression:=Current_Percentage, Format:="Percent")
            
            'Updating the Summary Table with the yearly change and the corresponding percentage
            Range("J" & Summary_Table_Row).Value = Current_Close - Current_Open
            Range("K" & Summary_Table_Row).Value = Current_Percentage_formatted
              
            'Color coding based on whether it had a positive or negative close relative to the open. Green (50) for positive, red (30 for negative.
            If Current_Percentage > 0 Then
            
                'Green if greater than 0
               Cells(Summary_Table_Row, 10).Interior.ColorIndex = 50
               
            Else
                'Red if less than
                 Cells(Summary_Table_Row, 10).Interior.ColorIndex = 30
            End If
                                   
            'That row is complete, so 1 is added to the Summary_Table_Row to get ready for the next new stock.
            Summary_Table_Row = Summary_Table_Row + 1
            
        End If
    End If
Next I

    'Setting up table for Bonus Section
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
'Finding the lastrow in the summary section
LastRowSummary = Cells(Rows.Count, "I").End(xlUp).Row

'Determining the min and max for Greatest % gain and loss as well as the greatest folume
Range("P2") = Format(Expression:=Application.WorksheetFunction.Max(Range("K1:K" & LastRowSummary), 1), Format:="Percent")
Range("P3") = Format(Expression:=Application.WorksheetFunction.Min(Range("K1:K" & LastRowSummary), 1), Format:="Percent")
Range("P4") = Application.WorksheetFunction.Max(Range("K1:K" & LastRowSummary))

'Finding the ticker symbol associated with the Greatest Values in the Bonus Table
For m = 2 To Summary_Table_Row
    'If the cell in the summary table equals the Greatest value in the Bonus Table, then populate the ticker next to that value.
    
    'Does this stock have the greatest increase then add it to the bonus table
    If Cells(m, 11).Value = Cells(2, 16).Value Then
        Cells(2, 15).Value = Cells(m, 9).Value
        
    'If not, does it have the greatest decres
    ElseIf Cells(m, 11).Value = Cells(3, 16).Value Then
        Cells(3, 15).Value = Cells(m, 9).Value
        
    'Then end
    End If
    
    'Does this stock have the greated volume, then add it to the bonus table.
    If Cells(m, 12).Value = Cells(4, 16).Value Then
        Cells(4, 15).Value = Cells(m, 9).Value
    End If
Next m
End Sub

   
'******* DONE!!!! Woot!!*********




