Sub challenge():

    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' Ticker Symbol:
        
        ' Set a header for the Ticker column
        ws.Cells(1, 9).Value = "Ticker"
        ' set a variable that holds the ticker symbol
        Dim Ticker As String
        ' Location of the Ticker
        Dim Ticker_Row As Integer
        Ticker_Row = 2
        
        ' Yearly Change:
        
        ' Set a header for the Yearly Change column
        ws.Cells(1, 10).Value = "Yearly Change"
        ' Set a variable for Closing_Price & Opening_Price & Yearly_Change
        Dim Closing_Price As Double
        Dim Opening_Price As Double
        ' Initializing the Opening Price value
        Opening_Price = ws.Cells(2, 3).Value
        Dim Yearly_Change As Double
        ' Location of the Yearly_Change
        Dim Yearly_Change_Row As Integer
        Yearly_Change_Row = 2
        
        ' Percent Change:
        
        ' Set a header for the Percent Change column
        ws.Cells(1, 11).Value = "Percent Change"
        ' Set a variable for Percent Change
        Dim Percent_Change As Double
        ' Location of the Percent_Change
        Dim Percent_Change_Row As Integer
        Percent_Change_Row = 2
        
        'Total stock volume:
        
        ' set a header for the Total Stock Volume column
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Set a variable for Total Stock Volume
        Dim Total_Stock_Volume As Double
        ' Initializing the Total_Stock_Volume value
        Total_Stock_Volume = 0
        ' Location of the Total_Stock_Volume
        Dim Total_Stock_Volume_Row As Integer
        Total_Stock_Volume_Row = 2
        
        ' Last part:
        
        ' Set the header for the last part:
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Set a header for Greatest % Increase
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ' Set a variable for Greatest % Increase
        Dim Max_Value As Double
        ' Initializing the Max value
        Max_Value = ws.Cells(2, 11).Value
        
        ' Set a header for Greates % Decrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ' Set a variable for Greatest % Deacrese
        Dim Min As Double
        ' Initializing the Min value
        Min_Value = ws.Cells(2, 11).Value
        
        ' Set a header for Greatest Total Volume
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ' Set a variable for Greates Total Volume
        Dim Greatest_Total_Volume As Double
        ' Initializing the Greatest_Total_Volume
        Greatest_Total_Volume = ws.Cells(2, 12).Value
        
        ' Find the last row:
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Find the last row for the last part:
        LastRow_2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        
        
        
        ' Loop through all data:
        For i = 2 To LastRow
            ' Checks to see if the Tickers are still the same if not, it displays it on the Ticker column and goes through the next one and repeats the whole process again.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Ticker_Row).Value = Ticker
               ' It takes the closing price of each ticker and takes the opening price of each ticker as it goes to the next one.
                Closing_Price = ws.Cells(i, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                ws.Range("J" & Yearly_Change_Row).Value = Yearly_Change
                    
                    ' Highlights positive changes with green and negative changes & no changes to red
                    If Yearly_Change > 0 Then
                        ws.Range("J" & Yearly_Change_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Yearly_Change_Row).Interior.ColorIndex = 3
                    End If
                 
                 ' In order to find the percent change, we need to divide the yearly change by the opening price and multiply it by 100. This repeats for each ticker.
                 Percent_Change = (Yearly_Change / Opening_Price) * 100
                 Opening_Price = ws.Cells(i + 1, 3).Value
                 ws.Range("K" & Percent_Change_Row).Value = Percent_Change & "%"
                 
                 ' Highlights positive changes with green and negative changes & no changes to red
                    If Percent_Change > 0 Then
                        ws.Range("K" & Percent_Change_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("K" & Percent_Change_Row).Interior.ColorIndex = 3
                    End If
                 
                 ' The total stock volume is the sum of all the volumes for each ticker.
                 Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                 ws.Range("L" & Total_Stock_Volume_Row).Value = Total_Stock_Volume
                 
                
                ' This just simply changes the rows and goes to the next one.
                Ticker_Row = Ticker_Row + 1
                Yearly_Change_Row = Yearly_Change_Row + 1
                Percent_Change_Row = Percent_Change_Row + 1
                Total_Stock_Volume_Row = Total_Stock_Volume_Row + 1
                
                ' Resetting the Total_Stock_Volume value
                Total_Stock_Volume = 0
                
            Else
            
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
            End If
       
        Next i
        
        ' Loop through all data for the last part
        For j = 2 To LastRow_2
            
            ' Finding the greatest % Increase
            If ws.Cells(j, 11).Value >= Max_Value Then
                Max_Value = ws.Cells(j, 11).Value
                ws.Cells(2, 17).Value = (Max_Value * 100) & "%"
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            End If
            
            ' Finding the greatest % Decrease
            If ws.Cells(j, 11).Value <= Min_Value Then
                Min_Value = ws.Cells(j, 11).Value
                ws.Cells(3, 17).Value = (Min_Value * 100) & "%"
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            End If
            
            ' Finding the greatest total volume
            If ws.Cells(j, 12).Value >= Greatest_Total_Volume Then
                Greatest_Total_Volume = ws.Cells(j, 12).Value
                ws.Cells(4, 17).Value = Greatest_Total_Volume
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            End If
        
        Next j
        
    Next ws

End Sub

