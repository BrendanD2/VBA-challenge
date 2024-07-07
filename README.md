# VBA-challenge

Code

Sub stock_challenge()
    'Setting up a loop to go through each worksheet
    For Each ws In Worksheets
        'Creating the header names for the values that we will find for each ticker.
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Finding the last row of the worksheet and setting it equal to a dimension so we can loop through it later.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Creating dimensions for the ticker, first open price, last closing price, and volume as they will help us find the desired values we need later.
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim vol As Double
        vol = 0
        'Creating a position dimension to make is easier to input our data that we are finding.
        Dim position As Integer
        position = 2
        
        'Going to loop through every row until the bottom of the worksheet.
        For i = 2 To LastRow
            'Creating an if-statement for when the ticker value changes, This means that we have looped through a ticker and can enter the desired values we need.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Setting the ticker equal to the value in the last spot given it is the last one on the worksheet.
                ticker = ws.Cells(i, 1).Value
                
                'Setting the ticker value for our new list equal to what we found in the I column where we are placing the ticker name.
                ws.Range("I" & position).Value = ticker
                
                'Last closing price is found using the data from the last cell where the closing price is marked.
                close_price = ws.Cells(i, 6).Value
                
                'Quarterly change is found by subtracting the first open price from the last closing price.
                quart_change = close_price - open_price
                
                'Making our final addition to the total volume by adding in the last value from the row.
                vol = vol + ws.Cells(i, 7).Value
                
                'Placing the quarterly change in the desired column position where we are tracking quarterly change.
                ws.Range("J" & position).Value = quart_change
                
                'Creating an If-Elseif statement for conditional formatting when the quarterly change is positive (green) or negative (red).
                If quart_change > 0 Then
                    ws.Range("J" & position).Interior.Color = vbGreen
                ElseIf quart_change < 0 Then
                    ws.Range("J" & position).Interior.Color = vbRed
                End If
                
                'Finding the Percent Change by dividing the first open price from the difference of the first open price from the last closing price.
                percent_chng = (close_price - open_price) / open_price
                
                'Setting the percent change value in the desired column position.
                ws.Range("K" & position).Value = percent_chng
                'Using Conditional formatting to make the percent change value a percentage.
                ws.Range("K" & position).NumberFormat = "0.00%"
                'Setting the total volume to the desired column position
                ws.Range("L" & position).Value = vol
                
                'Adding 1 to the position value so that the next entry does not enter on top of the last placement as well as resetting our vol (total volume) variable back to 0.
                position = position + 1
                vol = 0
                
            'Creating an else statement to continue adding to the total volume from all of the rows inbetween the first and last.
            Else
                'Creating an if statement for when the prior rows ticker value is different to find the first open price of the ticker.
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    open_price = ws.Cells(i, 3).Value
                    End If
                'Adding to volume based on the value from that row.
                vol = vol + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'After loop goes through every value we are finding the last row from the new data we created for each ticker.
        Last_newRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Creating dimensions to represent the Greatest percent increase and decrease with the ticker for those values, and the greatest total volume and the ticker for that value.
        Dim Grt_pctinc As Double
        Dim Grt_pctdec As Double
        Dim Grt_ttlvol As Double
        Dim Grt_inctick As String
        Dim Grt_dectick As String
        Dim Grt_voltick As String
        
        'Setting the variables tracking the greatest percentage increase, decrease, and total volume to 0.
        Grt_pctinc = 0
        Grt_pctdec = 0
        Grt_ttlvol = 0
        
        'Creating headers for the new data we are about to find.
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Looping through the new found data set.
        For k = 2 To Last_newRow
            'Finding the greatest percentage increase by testing the value from that cell on the Grt_pctinc variabel.
            If ws.Cells(k, 11).Value > Grt_pctinc Then
                'If the cells value is greater than the Grt_pctinc variable then it takes the new spot for greatest percentage increase and the ticker name is also recorded.
                Grt_pctinc = ws.Cells(k, 11).Value
                Grt_inctick = ws.Cells(k, 9).Value
                End If
        Next k
        
        'Finding the greatest percentage decrease by testing the value from that cell on the Grt_pctdec variabel.
        For j = 2 To Last_newRow
            If ws.Cells(j, 11).Value < Grt_pctdec Then
                'If the cells value is less than the Grt_pctdec variable then it takes the new spot for greatest percentage decrease and the ticker name is also recorded.
                Grt_pctdec = ws.Cells(j, 11).Value
                Grt_dectick = ws.Cells(j, 9).Value
                End If
        Next j
        
        'Finding the greatest total volume from the new data.
        For y = 2 To Last_newRow
            'If the cell value with toal volume is greater than the Grt_ttlvol variable then that variable is changed to the value of that cell and the ticker name is recorded.
            If ws.Cells(y, 12).Value > Grt_ttlvol Then
                Grt_ttlvol = ws.Cells(y, 12).Value
                Grt_voltick = ws.Cells(y, 9).Value
                End If
        Next y
        
        'After looping through and finding the values we then set the desired cells the values we found and format the percentage values.
        ws.Range("P2").Value = Grt_inctick
        ws.Range("Q2").Value = Grt_pctinc
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("P3").Value = Grt_dectick
        ws.Range("Q3").Value = Grt_pctdec
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("P4").Value = Grt_voltick
        ws.Range("Q4").Value = Grt_ttlvol
        
    Next ws
                
        
End Sub
