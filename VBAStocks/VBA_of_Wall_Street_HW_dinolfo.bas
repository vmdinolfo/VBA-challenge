Attribute VB_Name = "Module1"
Sub Stock_Test():

For Each ws In Worksheets
    '--------------------------------------------
    ' Set column headers for Summary Table
    '--------------------------------------------
    ws.Cells(1, 10) = "Ticker"
    ws.Cells(1, 11) = "Yearly Change"
    ws.Cells(1, 12) = "Percent Change"
    ws.Cells(1, 13) = "Total Stock Volume"
    ws.Cells(2, 16) = "Greatest % Increase"
    ws.Cells(3, 16) = "Greatest % Decrease"
    ws.Cells(4, 16) = "Greatest Total volume"
    ws.Cells(1, 17) = "Ticker"
    ws.Cells(1, 18) = "Value"
    
    '--------------------------------------------
    ' Set Color Index Variables
    '--------------------------------------------
    ColorRed = 3
    ColorGreen = 4
    
    '--------------------------------------------
    ' Set initial variable to hold Ticker Symbol
    '--------------------------------------------
    Dim Ticker_Symbol As String
    
    '--------------------------------------------
    ' Set initial variable for holding total stock volume per Ticker Symbol
    '--------------------------------------------
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    '--------------------------------------------
    ' Set Row Start for Yearly Change & Percent Change
    '--------------------------------------------
    Dim Row_Start As Double
    Row_Start = 2
    '--------------------------------------------
    ' Keep track of the location for each ticker symbol in the Summary Table
    '--------------------------------------------
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    '--------------------------------------------
    ' Loop through all stocks
    '--------------------------------------------
    row_length = ws.Range("A1").End(xlDown).Row                                     ' Got bottom of table row length with help from TA Chris
    For t = 2 To row_length
        
        If ws.Cells(t + 1, 1).Value <> ws.Cells(t, 1).Value Then
        
       '--------------------------------------------
       ' Set the Ticker Symbol
       '--------------------------------------------
            Ticker_Symbol = ws.Cells(t, 1)
            
       '--------------------------------------------
       ' Print the Ticker Symbol in the Summary Table
       '--------------------------------------------
            ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
            
       '--------------------------------------------
       ' Sum the Total Stock Volume for Ticker Symbol
       '--------------------------------------------
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(t, 7).Value
            
       '--------------------------------------------
       ' Print the Total Stock Volume in the Summary Table
       '--------------------------------------------
           ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
           
       '--------------------------------------------
       ' Calculate Yearly Change for Ticker Symbol
       '--------------------------------------------
            Dim Annual_Open As Double
            Annual_Open = ws.Cells(Row_Start, 3).Value
            
            Dim Yearly_Change As Double
            Yearly_Change = ws.Cells(t, 6).Value - Annual_Open
            
       '--------------------------------------------
       ' Print Yearly Change for Ticker Symbol in the Summary Table
       '--------------------------------------------
            ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
            If Yearly_Change < 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = ColorRed
            Else
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = ColorGreen
            End If
       '--------------------------------------------
       ' Calculate Percent Change for Ticker Symbol
       '--------------------------------------------
        
            Dim Percent_Change As Double
            If Annual_Open = 0 Then
                Annual_Open = ws.Cells(Row_Start + 1, 3)
            Else
                Percent_Change = Yearly_Change / Annual_Open
            End If
            
       '--------------------------------------------
       ' Print Percent Change for Ticker Symbol in the Summary Table
       '--------------------------------------------
            ws.Range("L" & Summary_Table_Row).Value = Percent_Change
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
       '--------------------------------------------
       ' Reset Total Stock Volume
       '--------------------------------------------
            Total_Stock_Volume = 0
               
       '--------------------------------------------
       ' Add one row to the Summary Table
       '--------------------------------------------
            Summary_Table_Row = Summary_Table_Row + 1
       '--------------------------------------------
       ' Reset Row_Start
       '--------------------------------------------
            Row_Start = t + 1
    '--------------------------------------------
    ' Reset
    '--------------------------------------------
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(t, 7).Value
            
        End If

   Next t
    '--------------------------------------------
    ' Find bottom row# of Summary Table
    '--------------------------------------------
    row_length2 = ws.Range("J1").End(xlDown).Row                                    ' Got bottom of table row length with help from TA Chris
    '--------------------------------------------
    ' Find & Print Greatest % Increase
    '--------------------------------------------
    Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("L:L"))          ' Equation for Max value from https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba/42633375
    For i = 2 To row_length2
        If ws.Cells(i, 12).Value = Greatest_Increase Then
            Ticker_Greatest_Increase = ws.Cells(i, 10).Value
        End If
     Next i
    ws.Cells(2, 17).Value = Ticker_Greatest_Increase
    ws.Cells(2, 18).Value = Greatest_Increase
    ws.Cells(2, 18).NumberFormat = "0.00%"
    '--------------------------------------------
    ' Find & Print Greatest % Decrease
    '--------------------------------------------
    Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("L:L"))
    For j = 2 To row_length2
        If ws.Cells(j, 12).Value = Greatest_Decrease Then
            Ticker_Greatest_Decrease = ws.Cells(j, 10).Value
        End If
     Next j
    ws.Cells(3, 17).Value = Ticker_Greatest_Decrease
    ws.Cells(3, 18).Value = Greatest_Decrease
    ws.Cells(3, 18).NumberFormat = "0.00%"
    '--------------------------------------------
    ' Find & Print Greatest Total Volume
    '--------------------------------------------
    Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Range("M:M"))
    For x = 2 To row_length2
        If ws.Cells(x, 13).Value = Greatest_Total_Volume Then
            Ticker_Greatest_Total_Volume = ws.Cells(x, 10).Value
        End If
    Next x
    ws.Cells(4, 17).Value = Ticker_Greatest_Total_Volume
    ws.Cells(4, 18).Value = Greatest_Total_Volume
    ws.Cells(4, 18).NumberFormat = "0"
    
    
Next ws

MsgBox ("All Done!")
    
End Sub
