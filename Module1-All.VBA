Attribute VB_Name = "Module1"
'Instructions: By adding loop for Worksheest I went through them
'Create a script that loops through all the stocks for each quarter and outputs the following information:
'The ticker symbol
'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
'The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
'The total stock volume of the stock.

Sub Stockanalysis()

'Loop for all the sheets.
'Set a variables

    For Each ws In Worksheets

    Dim ticker As String                  'for ticker
    Dim volume As Double                'for vol
        volume = 0
    Dim open_price As Double
    Dim close_price As Double           'for calculating change
'Quarterly Change:(Close Price - Open Price)
           
    Dim Quarterly_change As Double
      
'Percent((Close - Open)/Open)*100
    Dim percent_change As Double        'for having percent
    
            
'to identify the first cell to put result(summary)
        Dim summary As Integer
        summary = 2
      
'cell that the open price has start
        open_price = ws.Cells(2, 3).Value



'Title for the Summary headers as in ReadMe

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'Count the number of rows in the first column.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'start the Loop on ticker (Column A)

        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              ticker = ws.Cells(i, 1).Value

              'Set the volume
              volume = volume + ws.Cells(i, 7).Value

              'Print the ticker name in the summary (Column I)
              ws.Range("I" & summary).Value = ticker

              'Print the volume for each ticker in the summary (Column L)
              ws.Range("L" & summary).Value = volume

'We have to identify close_price here in loop, if we give pair location like (2,6)it wouldn;t be in the loop to go to the next row
              close_price = ws.Cells(i, 6).Value

              'Quarterly change
               Quarterly_change = (close_price - open_price)
              
              'Print the quarterly change for each unique ticker name in the summary (Column J)
              ws.Range("J" & summary).Value = Quarterly_change
'if wd devide by 0 there will be error, there are some rows that change is 0; so we need condition

                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = Quarterly_change / open_price
                
                End If

              'Print the quarterly change for each ticker name in the summary (Column J)
              ws.Range("K" & summary).Value = percent_change
              ws.Range("K" & summary).NumberFormat = "0.00%"
   
              'now we need to go to the next row in summary
              summary = summary + 1

              'Reset volume of trade to zero, to start add up for new ticker
              volume = 0

              'go to next row to grab new open price
              open_price = ws.Cells(i + 1, 3)
'And again check the value of the column
            Else
              
               'Add the volume
              volume = volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

'To highlight positive change in green and negative change in red with Conditional formatting, again we need to know the last row (this time for Quarterly Change /Column J)
'This part can be out of the loop and run separetly

    lastrow_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code quarterly change
        For i = 2 To lastrow_summary
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If

        
        Next i

'To have information (Statisticall measurments)as second photo in ReadMe, we need to print names in certain cells

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'To be able to checkwhole column J we go through the loop
        For i = 2 To lastrow_summary
        
            'Find the maximum percent change
            If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
'I googled percentage since my numbers show as flout with decimals first

                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub

